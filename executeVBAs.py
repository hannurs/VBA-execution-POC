import win32com.client
from win32com.client import gencache
import pywintypes
import pythoncom
import os, os.path
import sys
import shutil
from azure.storage.blob import BlobServiceClient
import time
from typing import List

CONNECTION_STRING = "blob_connection_string"

def ListVBAMacros(input_file: str):
    """List all VBA macros (Subroutines/Functions) in the Excel workbook"""
    
    # Open Excel application
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False  # Keep Excel hidden
    macros = []
    log_file = open("log.txt", "a")
    log_file.write("Excel App Created\n")

    try:
        log_file.write(input_file)
        workbook = excel_app.Workbooks.Open(input_file, ReadOnly=1)
        log_file.write(f"{input_file} opened without error")
        vb_project = workbook.VBProject
        for vb_component in vb_project.VBComponents:
            code_module = vb_component.CodeModule
            num_lines = code_module.CountOfLines
            
            # Loop through each line of code to find the macros
            for line_num in range(1, num_lines + 1):
                line = code_module.Lines(line_num, 1).strip()

                # Identify macro definitions (starting with "Sub" or "Function")
                if line.startswith("Sub ") or line.startswith("Function "):
                    macro_name = line.split()[1].split("(")[0]
                    macros.append(macro_name)
                    print(f"Found macro: {macro_name}")
        with open("log.txt", "a") as log_file:
            log_file.write("All good with ListVBAMacros\n")

    except pywintypes.com_error as e:
        # Print details of the COM error
        with open("log.txt", "a") as log_file:
            log_file.write("COM Error:")
            log_file.write(f"  HRESULT: {hex(e.hresult)}")
            log_file.write(f"  Source: {e.strerror}")
            log_file.write(f"  Description: {e.excepinfo[2] if e.excepinfo[2] else 'No description available'}\n")
            log_file.write(f"  Helpfile: {e.excepinfo[3] if e.excepinfo[3] else 'No helpfile available'}\n")

    except Exception as e:
        with open("log.txt", "a") as log_file:
            log_file.write("Test\n")
            log_file.write(str(type(e)))
            # log_file.write(f"Error occurred:\n{traceback.format_exc()}\n")
            # log_file.write(f"Error: {e}")
        print(f"Error: {str(e)}")
    
    finally:
        # Close the workbook and Excel application
        workbook.Close(False)
        excel_app.Quit()
        return macros

def ConnectToBlobService() -> BlobServiceClient:
    return BlobServiceClient.from_connection_string(CONNECTION_STRING)


def ListBlobsInContainer(
        blob_service_client: BlobServiceClient,
        container_name: str,
        ):
    return blob_service_client.get_container_client(container=container_name).list_blobs()


def ListBlobNamesInContainerAsList(
        blob_service_client: BlobServiceClient,
        container_name: str,
        ):
    blobs_list = []
    blobs_paged = blob_service_client.get_container_client(container=container_name).list_blob_names()
    for blob in blobs_paged:
        blobs_list.append(blob)

    return blobs_list


def DownloadBlob(
        blob_service_client: BlobServiceClient,
        container_name: str,
        blob_name: str
) -> str:
    blob_client = blob_service_client.get_container_client(container=container_name).get_blob_client(blob=blob_name)
    download_dir = os.path.join(os.getcwd(), container_name)
    download_path = os.path.join(download_dir, blob_name)
    with open(download_path, "wb") as download_file:
        download_file.write(blob_client.download_blob().readall())
    print(f"Downloaded {blob_name} to {download_path}")
    return download_path


def UploadBlob(
        blob_service_client: BlobServiceClient,
        container_name: str,
        blob_name: str
) -> None:
    blob_client = blob_service_client.get_container_client(container=container_name).get_blob_client(blob_name)
    try:
        # Upload the file
        local_file_path = os.path.join(os.getcwd(), container_name, blob_name)
        with open(local_file_path, "rb") as data:
            blob_client.upload_blob(data, overwrite=True)
        
        print(f"File {local_file_path} uploaded to {blob_name} in container {container_name}.")
    
    except Exception as e:
        print(f"Error uploading blob: {str(e)}")
    return

def ListFilesInDirectory(directory):
    shell = win32com.client.Dispatch("Scripting.FileSystemObject")
    folder = shell.GetFolder(directory)
    
    file_list = []
    
    # Iterate over all files in the folder
    for file in folder.Files:
        file_list.append(file.Name)

    return file_list


# This function allows us to run Excel headlessly
# Without this, the below code will open a visible Excel instance although the process will still work

def EnsureDispatchEx(clsid, new_instance=True):
    """Create a new COM instance and ensure cache is built,
       unset read-only gencache flag"""
    if new_instance:
        clsid = pythoncom.CoCreateInstanceEx(clsid, None, pythoncom.CLSCTX_SERVER,
                                             None, (pythoncom.IID_IDispatch,))[0]
    if gencache.is_readonly:
        #fix for "freezed" app: py2exe.org/index.cgi/UsingEnsureDispatch
        gencache.is_readonly = False
        gencache.Rebuild()
    try:
        return gencache.EnsureDispatch(clsid)
    except (KeyError, AttributeError):  # no attribute 'CLSIDToClassMap'
        # something went wrong, reset cache
        shutil.rmtree(gencache.GetGeneratePath())
        for i in [i for i in sys.modules if i.startswith("win32com.gen_py.")]:
            del sys.modules[i]
        return gencache.EnsureDispatch(clsid)

def Run(input_file: str, output_file: str, vba_script: str, param: str = None):
    """Function to run a Excel spreadsheet, execute a VBA script and save the output
    """
    xl = EnsureDispatchEx("Excel.Application") 
    # win32com.client.Dispatch("Excel.Application")
    # Without using the custom function above

    wb = xl.Workbooks.Open(input_file, ReadOnly=1)
    xl.Run(vba_script, param)
    if os.path.exists(output_file):
        os.remove(output_file)
    wb.SaveAs(output_file)
    xl.Quit()


def DownloadBlobsFromContainer(
        container_name: str,
        blob_list: List[str]
        ) -> None:
    blob_service_client = ConnectToBlobService()
    print(f"Searching for new files in {container_name} container in blob storage...")
    os.mkdir(os.path.join(os.getcwd(), container_name))
    for blob in blob_list:
        DownloadBlob(blob_service_client=blob_service_client, container_name=container_name, blob_name=blob)


def ExecuteVBAs(
        input_dir_name: str,
        output_dir_name: str,
        param: str = None
        ) -> None:
    input_dir_path = os.path.join(os.getcwd(), input_dir_name)
    os.mkdir(os.path.join(os.getcwd(), output_dir_name))
    output_dir_path = os.path.join(os.getcwd(), output_dir_name)
    input_files = ListFilesInDirectory(input_dir_path)
    for input_file in input_files:
        full_input_path = os.path.join(input_dir_path, input_file)
        full_output_path = os.path.join(output_dir_path, input_file)
        print("Processing", input_file, "file")
        macros = ListVBAMacros(full_input_path)
        for macro in macros:
            Run(full_input_path, full_output_path, input_file + "!" + macro, param=param)
    return


def UploadBlobsToContainer(container_name: str) -> None:
    blob_service_client = ConnectToBlobService()
    blobs = ListBlobNamesInContainerAsList(blob_service_client=blob_service_client, container_name=container_name)
    outputs_dir = os.path.join(os.getcwd(), container_name)
    output_files = ListFilesInDirectory(outputs_dir)
    for file in output_files:
        if (file not in blobs):
            UploadBlob(blob_service_client=blob_service_client, container_name=container_name, blob_name=file)


def ClearLocalFolders(folders: List[str]):
    for folder in folders:
        shutil.rmtree(os.path.join(os.getcwd(), folder))

def CheckNewFilesToProcess(
        input_container: str,
        output_container: str
) -> List[str]:
    blob_service_client = ConnectToBlobService()
    input_files = ListBlobNamesInContainerAsList(blob_service_client=blob_service_client, container_name=input_container)
    output_files = ListBlobNamesInContainerAsList(blob_service_client=blob_service_client, container_name=output_container)
    new_files = []
    for file in input_files:
        if (file not in output_files):
            new_files.append(file)

    return new_files

def main():
    input_container_name = "input"
    output_container_name = "output"

    while(True):
        blobs_to_process = CheckNewFilesToProcess(input_container=input_container_name, output_container=output_container_name)
        DownloadBlobsFromContainer(container_name=input_container_name, blob_list=blobs_to_process)
        ExecuteVBAs(input_dir_name=input_container_name, output_dir_name=output_container_name)
        UploadBlobsToContainer(container_name=output_container_name)
        ClearLocalFolders([input_container_name, output_container_name])
        time.sleep(5)

    return 1

if __name__ == '__main__':
    main()
