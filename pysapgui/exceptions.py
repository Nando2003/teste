from typing import Optional


class UnableToConnectException(Exception):
    """Exception raised when unable to connect to the SAP GUI."""
    def __init__(self):
        message = (
            "Unable to connect to the SAP GUI. "
            "Please ensure that the SAP GUI is installed and running, "
            "and that the scripting API is enabled in the SAP GUI options."
        )
        super().__init__(message)


class NoSapConnectionException(Exception):
    """Exception raised when no SAP connection is found."""
    def __init__(self, connection_id: Optional[int] = None):
        if connection_id:
            message = (
                f"No SAP connection with ID '{connection_id}' found. "
                "Please ensure that you have an active SAP connection."
            )
        else:
            message = (
                "No SAP connection found. "
                "Please ensure that you have an active SAP connection."
            ) 
        super().__init__(message)
        

class NoSapSessionException(Exception):
    """Exception raised when no SAP session is found."""
    def __init__(self, session_id: Optional[int] = None, is_busy: bool = False):
        if session_id:
            message = (
                f"No SAP session with ID '{session_id}' found. "
                "Please ensure that you have an active SAP session."
            )
        else:
            message = (
                "No SAP session found. "
                "Please ensure that you have an active SAP session."
            )
            
        if is_busy:
            message += " The session is currently busy and cannot be accessed."
            
        super().__init__(message)
        

class SapElementNotFoundException(Exception):
    """Exception raised when a SAP element is not found."""
    def __init__(self, element_id: str):
        message = f"SAP element with ID '{element_id}' not found."
        super().__init__(message)


class SapSessionMismatchException(Exception):
    """Exception raised when there is a mismatch in the SAP session."""
    def __init__(self):
        message = (
            "The SAP session has changed. "
            "Please refresh the session to continue."
        )
        super().__init__(message)
        
        
class SapAttributeNotFoundException(AttributeError):
    """Exception raised when a SAP element attribute is not found."""
    def __init__(self, attribute: str):
        message = f"SAP element has no attribute '{attribute}'."
        super().__init__(message)