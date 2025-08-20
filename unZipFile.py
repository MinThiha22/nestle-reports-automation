import zipfile
import os
from enum import Enum

class ExtractionResult:
    def __init__(self, success=False, extract_path=None, message="", error_type=None):
        self.success = success
        self.extract_path = extract_path
        self.message = message
        self.error_type = error_type
        
class ExtractionError(Enum):
    INVALID_ZIP = "invalid_zip"
    PERMISSION_ERROR = "permission_error"
    DISK_SPACE_ERROR = "disk_space_error"
    GENERAL_ERROR = "general_error"   

def extract_zip(zip_path, extract_to):
    """
    Extracts a zip file to the specified directory.
    
    Args:
        zip_path (str): Path to the zip file
        extract_to (str): Directory to extract to
        
    Returns:
        ExtractionResult: Object containing extraction status and details
    """
    try:
        # Check if zip file exists
        if not os.path.exists(zip_path):
            return ExtractionResult(
                success=False,
                message=f"Zip file not found: {zip_path}",
                error_type=ExtractionError.GENERAL_ERROR
            )
        
        # Create extraction directory if it doesn't exist
        os.makedirs(extract_to, exist_ok=True)
        
        # Extract the zip file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        
        # Verify extraction was successful
        if os.path.exists(extract_to) and os.listdir(extract_to):
            return ExtractionResult(
                success=True,
                extract_path=extract_to,
                message=f"Successfully extracted to: {extract_to}"
            )
        else:
            return ExtractionResult(
                success=False,
                message=f"Extraction completed but no files found in: {extract_to}",
                error_type=ExtractionError.GENERAL_ERROR
            )
            
    except zipfile.BadZipFile:
        return ExtractionResult(
            success=False,
            message="The file is not a valid zip file or is corrupted",
            error_type=ExtractionError.INVALID_ZIP
        )
    except PermissionError:
        return ExtractionResult(
            success=False,
            message=f"Permission denied: Cannot write to {extract_to}",
            error_type=ExtractionError.PERMISSION_ERROR
        )
    except OSError as e:
        if "No space left on device" in str(e):
            return ExtractionResult(
                success=False,
                message="Insufficient disk space for extraction",
                error_type=ExtractionError.DISK_SPACE_ERROR
            )
        else:
            return ExtractionResult(
                success=False,
                message=f"OS Error during extraction: {e}",
                error_type=ExtractionError.GENERAL_ERROR
            )
    except Exception as e:
        return ExtractionResult(
            success=False,
            message=f"Unexpected error during extraction: {e}",
            error_type=ExtractionError.GENERAL_ERROR
        )



