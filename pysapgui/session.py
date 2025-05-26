import win32com.client as client
from typing import Optional, Any, TYPE_CHECKING

if TYPE_CHECKING:
    from pysapgui.element import Element

from pysapgui.connection import Connection
from pysapgui.utils import search_path
from pysapgui.exceptions import NoSapSessionException, SapElementNotFoundException


class Session:
    """
    Class to manage SAP GUI sessions.
    
    This class provides methods to connect to the SAP GUI, retrieve the current session,
    and refresh the session if needed.
    
    Attributes:
        session (client.CDispatch): The current SAP session.
        connection (Connection): The SAP connection object.
    """
    def __init__(
        self, 
        session_id: Optional[int] = None, 
        connection: Optional[Connection] = None
    ):
        self.connection = connection if connection else Connection()
        self.session = self.__get_session(session_id)
        self.session_id = session_id
    
    def __getattr__(self, name: Any) -> Any:
        return getattr(self.session, name)
    
    def __eq__(self, value: object) -> bool:
        return isinstance(value, Session) and self.get_id() == value.get_id()
    
    def __get_session(self, session_id: Optional[int]) -> client.CDispatch:
        sessions = self.connection.Sessions
        sessions_length = sessions.Count
        
        if not sessions_length:
            raise NoSapSessionException
        
        if session_id is not None:
            if (session_id < 0) or (sessions_length <= session_id):
                raise NoSapSessionException(session_id)
            
            session =sessions.Item(session_id)

            if not session.Busy:
                return session
            
            raise NoSapSessionException(session_id)
            
        for idx in range(sessions_length):
            session = sessions.Item(idx)
            if not session.Busy:
                return session
        
        raise NoSapSessionException(session_id)

    def get_id(self) -> str:
        """
        Get the ID of the current SAP session.
        
        Returns:
            str: The ID of the current SAP session.
        """
        return self.session.Id
    
    def refresh(self) -> None:
        """
        Refresh the current SAP session.
        
        This method retrieves the current session again to ensure it is up-to-date.
        """
        self.session = self.__get_session(self.session_id)
        self.connection.refresh()

    def goto_tcode(self, tcode: str) -> None:
        """
        Navigates to the specified transaction code in the SAP session.
        
        Args:
            tcode (str): The transaction code to navigate to.
        """        
        self.session.sendcommand(tcode)
    
    def send_vkey(self, vkey: str) -> None:
        """
        Sends a virtual key to the SAP session.
        
        Args:
            vkey (str): The virtual key to send.
        """
        self.session.findById("wnd[0]").sendVKey(vkey)
        
    def maximaze_window(self) -> None:
        """
        Maximizes the SAP GUI window.
        
        This method sets focus on the main window and maximizes it.
        """
        self._session.findById("wnd[0]").setFocus()
        self._session.findById("wnd[0]").maximize()
    
    def close_window(self) -> None:
        """
        Closes the current SAP GUI window.
        
        This method sets focus on the main window and closes it.
        """
        self._session.findById("wnd[0]").setFocus()
        self._session.findById("wnd[0]").close()
        
    def get_screen_region(self) -> tuple[int, int, int, int]:
        """
        Gets the screen region of the SAP GUI window.
        
        Returns:
            tuple[int, int, int, int]: A tuple containing the left, top, width, and height of the SAP GUI window.
        """
        window = self._session.findById("wnd[0]")
        return window.ScreenLeft, window.ScreenTop, window.width, window.height
    
    def get_window_title(self) -> str:
        """
        Gets the title of the current SAP GUI window.
        
        Returns:
            str: The title of the SAP GUI window.
        """
        return str(self._session.findById("wnd[0]").text)
    
    def find_element(self, element_id: str) -> 'Element':
        """
        Finds an element in the SAP GUI session by its ID.
        
        Args:
            element_id (str): The ID of the element to find.
        
        Returns:
            Any: The found SAP GUI element wrapped in an Element instance.
            
        Raises:
            SapElementNotFoundException: If the element with the specified ID is not found.
        """
        from pysapgui.element import Element
        try:
            return Element(self, self.session.findById(element_id))
        except Exception as e:
            raise SapElementNotFoundException(element_id) from e
    
    def find_partial_element(self, re_element_id: str) -> Optional['Element']:
        """
        Finds an element in the SAP GUI session by a partial ID.
        
        Args:
            re_element_id (str): The partial ID of the element to find.
        
        Returns:
            Optional[Element]: The found SAP GUI element wrapped in an Element instance, or None if not found.
        """
        from pysapgui.element import Element
        results = search_path(
            self.session,
            re_path=re_element_id,
            return_all=False,
            element_wrapper=lambda elem: Element(self, elem)
        )
        return results if isinstance(results, Element) else None
    