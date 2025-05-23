#include <ClrPhlib.h>
#include <UnmanagedPh.h>

using namespace Dependencies;
using namespace System::Text;
using namespace ClrPh;
using namespace Runtime::InteropServices;

PE::PE(
    _In_ String ^ Filepath
)
{
    this->m_Impl = new UnmanagedPE();
    this->Filepath = gcnew String(Filepath);
    this->LoadSuccessful = false;

    this->m_ExportsInit = false;
    this->m_ImportsInit = false;
}

PE::~PE()
{
    Unload();
    delete m_Impl;
}

PE::!PE() {
    Unload();
    delete m_Impl;
}

bool PE::Load()
{
    // Load PE as mapped section
    wchar_t* PvFilepath = (wchar_t*)(Marshal::StringToHGlobalUni(Filepath)).ToPointer();
    this->LoadSuccessful = m_Impl->LoadPE(PvFilepath);
	Marshal::FreeHGlobal(IntPtr((void*)PvFilepath));

	if (!LoadSuccessful) {
		return false;
	}
        
	// Parse PE
	LoadSuccessful &= InitProperties();
	if (!LoadSuccessful) {
		m_Impl->UnloadPE();
		return false;
	}


    return LoadSuccessful;
}

void PE::Unload()
{
    if (LoadSuccessful)
        m_Impl->UnloadPE();
}

bool PE::InitProperties()
{
    LARGE_INTEGER time;
    SYSTEMTIME systemTime;

    PH_MAPPED_IMAGE PvMappedImage = m_Impl->m_PvMappedImage;
    
    Properties = gcnew PeProperties();
    Properties->Machine = PvMappedImage.NtHeaders->FileHeader.Machine;
    Properties->Magic = m_Impl->m_PvMappedImage.Magic;
    Properties->Checksum = PvMappedImage.NtHeaders->OptionalHeader.CheckSum;
    Properties->CorrectChecksum = (Properties->Checksum == PhCheckSumMappedImage(&PvMappedImage));

    RtlSecondsSince1970ToTime(PvMappedImage.NtHeaders->FileHeader.TimeDateStamp, &time);
    PhLargeIntegerToLocalSystemTime(&systemTime, &time);
    Properties->Time = gcnew DateTime (systemTime.wYear, systemTime.wMonth, systemTime.wDay, systemTime.wHour, systemTime.wMinute, systemTime.wSecond, systemTime.wMilliseconds, DateTimeKind::Local);

    if (PvMappedImage.Magic == IMAGE_NT_OPTIONAL_HDR32_MAGIC)
    {
        PIMAGE_OPTIONAL_HEADER32 OptionalHeader = (PIMAGE_OPTIONAL_HEADER32) &PvMappedImage.NtHeaders->OptionalHeader;
        
        Properties->ImageBase = (Int64) OptionalHeader->ImageBase;
        Properties->SizeOfImage = OptionalHeader->SizeOfImage;
        Properties->EntryPoint = (Int64) OptionalHeader->AddressOfEntryPoint;
    }
    else
    {
        PIMAGE_OPTIONAL_HEADER64 OptionalHeader = (PIMAGE_OPTIONAL_HEADER64)&PvMappedImage.NtHeaders->OptionalHeader;

        Properties->ImageBase = (Int64)OptionalHeader->ImageBase;
        Properties->SizeOfImage = OptionalHeader->SizeOfImage;
        Properties->EntryPoint = (Int64)OptionalHeader->AddressOfEntryPoint;

    }

    Properties->Subsystem = PvMappedImage.NtHeaders->OptionalHeader.Subsystem;
    Properties->SubsystemVersion = gcnew Tuple<Int16, Int16>(
        PvMappedImage.NtHeaders->OptionalHeader.MajorSubsystemVersion,
        PvMappedImage.NtHeaders->OptionalHeader.MinorSubsystemVersion);
    Properties->Characteristics = PvMappedImage.NtHeaders->FileHeader.Characteristics;
    Properties->DllCharacteristics = PvMappedImage.NtHeaders->OptionalHeader.DllCharacteristics;

    Properties->FileSize = PvMappedImage.Size;
	return true;
}

Collections::Generic::List<PeExport^> ^ PE::GetExports()
{
    if (m_ExportsInit)
        return m_Exports;

    m_ExportsInit = true;
    m_Exports = gcnew Collections::Generic::List<PeExport^>();

    if (!LoadSuccessful)
        return m_Exports;

    if (NT_SUCCESS(PhGetMappedImageExports(&m_Impl->m_PvExports, &m_Impl->m_PvMappedImage)))
    {
        for (size_t Index = 0; Index < m_Impl->m_PvExports.NumberOfEntries; Index++)
        {
			PeExport^ exp = PeExport::FromMapimg(*m_Impl, Index);

			if (exp)
			{
				m_Exports->Add(exp);
			}

        }
    }

    return m_Exports;
}


Collections::Generic::List<PeImportDll^> ^ PE::GetImports()
{
    if (m_ImportsInit)
        return m_Imports;

    m_ImportsInit = true;
    m_Imports = gcnew Collections::Generic::List<PeImportDll^>();

    if (!LoadSuccessful)
        return m_Imports;

    // Standard Imports
    if (NT_SUCCESS(PhGetMappedImageImports(&m_Impl->m_PvImports, &m_Impl->m_PvMappedImage)))
    {
        for (size_t IndexDll = 0; IndexDll< m_Impl->m_PvImports.NumberOfDlls; IndexDll++)
        {
            m_Imports->Add(gcnew PeImportDll(&m_Impl->m_PvImports, IndexDll));
        }
    }

    // Delayed Imports
    if (NT_SUCCESS(PhGetMappedImageDelayImports(&m_Impl->m_PvDelayImports, &m_Impl->m_PvMappedImage)))
    {
        for (size_t IndexDll = 0; IndexDll< m_Impl->m_PvDelayImports.NumberOfDlls; IndexDll++)
        {
            m_Imports->Add(gcnew PeImportDll(&m_Impl->m_PvDelayImports, IndexDll));
        }
    }

    return m_Imports;
}



String^ PE::GetManifest()
{
    if (!LoadSuccessful)
        return gcnew String("");

    // Extract embedded manifest
    INT  rawManifestLen;
    BYTE* rawManifest;
    if (!m_Impl->GetPeManifest(&rawManifest, &rawManifestLen))
        return gcnew String("");


    // Converting to wchar* and passing it to a C#-recognized String object
    UTF8Encoding Utf8Decoder;

    array<unsigned char> ^buffer = gcnew array<unsigned char>(rawManifestLen + 1);
    for (int i = 0; i < rawManifestLen; i++)
    {
        buffer[i] = rawManifest[i];
    }
    buffer[rawManifestLen] = 0;

    return  Utf8Decoder.GetString(buffer, 0, rawManifestLen);
}

bool PE::IsWow64Dll() 
{
    return ((Properties->Machine & 0xffff ) == IMAGE_FILE_MACHINE_I386);
}

bool PE::IsArm32Dll()
{
  return ((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_ARMNT);
}

bool PE::IsClrDll()
{
    PIMAGE_DATA_DIRECTORY dataDirectory;
    if (NT_SUCCESS(PhGetMappedImageDataEntry( &m_Impl->m_PvMappedImage, IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR, &dataDirectory)))
    {
        return dataDirectory->VirtualAddress != 0;
    }
    return false;
}

String^ PE::GetProcessor()
{
  if ((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_I386)
    return gcnew String("x86");
  if ((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_ARMNT)
    return gcnew String("arm");
  if ((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_ARM64)
    return gcnew String("arm64");
  if ((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_AMD64)
    return gcnew String("amd64");
  
  return gcnew String("unknown");
}

bool PE::CheckProcessor(String^ ProcessorArch)
{
  if (((Properties->Machine & 0xffff) == IMAGE_FILE_MACHINE_ARM64) && (ProcessorArch == "amd64"))
    return true;

  if (GetProcessor() == ProcessorArch)
    return true;
  return false;
}
