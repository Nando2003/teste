import win32com.client as client
from typing import Optional, Any, TYPE_CHECKING

if TYPE_CHECKING:
    from pysapgui.session import Session

from pysapgui.exceptions import UnableToConnectException, NoSapConnectionException


class Connection:
    def __init__(self, connection_id: Optional[int] = None):
        self.ScriptingEngine = self.__get_scripting_engine()
        self.connection = self.__get_connection(connection_id)
        self.connection_id = connection_id
    
    def __getattr__(self, name: Any) -> Any:
        return getattr(self.session, name)
    
    def __eq__(self, value: object) -> bool:
        return isinstance(value, Connection) and self.get_id() == value.get_id()
    
    def __get_scripting_engine(self) -> client.CDispatch:
        try:
            SapGuiAuto = client.GetObject("SAPGUI")
            return SapGuiAuto.GetScriptingEngine
        except Exception as e:
            raise UnableToConnectException from e
        
    def __get_connection(self, connection_id: Optional[int] = None) -> client.CDispatch:
        connections = self.ScriptingEngine.Connections
        connections_length = connections.Count
        
        if not connections_length:
            raise NoSapConnectionException
        
        if connection_id:
            if (connection_id < 0) or (connections_length < connection_id):
                raise NoSapConnectionException(connection_id)

            return connections.Item(connection_id)
        
        return connections.Item(connections_length - 1)
    
    def get_id(self) -> str:
        """
        Get the ID of the current SAP connection.
        
        Returns:
            str: The ID of the current SAP connection.
        """
        return self.connection.Id
    
    def refresh(self) -> None:
        """
        Refresh the current SAP connection.
        
        This method retrieves the current connection again to ensure it is up-to-date.
        """
        self.connection = self.__get_connection(self.connection_id)
    
    def get_session(self, session_id: Optional[int] = None) -> 'Session':
        """
        Get the current SAP session.
        
        Args:
            session_id (Optional[int]): The ID of the SAP session to retrieve. If None, retrieves the last active session.
        
        Returns:
            Session: An instance of the Session class representing the current SAP session.
        """
        from pysapgui.session import Session
        return Session(session_id, self)