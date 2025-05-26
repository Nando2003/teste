import re
import win32com.client as client
from typing import Optional, Any, Literal

from pysapgui.session import Session
from pysapgui.exceptions import SapSessionMismatchException, SapAttributeNotFoundException
from pysapgui.utils import search_path, search_element, requires_valid_session


class Element:
    def __init__(self, session: Session, element: client.CDispatch):
        self.session = session
        self.element = element

class SapElement:
    def __init__(self, session: Session, element: client.CDispatch):
        self.session = session
        self.element = element

        # Store the real session object for direct access
        self.__real_session = session.session
        
    def __getattr__(self, name: Any) -> Any:
        try:
            return getattr(self.element, name)
        except AttributeError as e:
            raise SapAttributeNotFoundException(name) from e
        
    def _verify_session(self) -> None:
        if self.session.session != self.__real_session:
            raise SapSessionMismatchException
        
    @requires_valid_session
    def click(self) -> None:
        """
        Simulates a click on the SAP element
        """
        self.element.press()
    
    def get_id(self) -> str:
        """
        Returns the ID of the SAP element
        """
        return self.element.Id
    
    def get_text(self) -> str:
        """
        Returns the title of the SAP element
        """
        return self.element.text
    
    @requires_valid_session
    def fill(self, value: Any) -> None:
        """
        Fills the SAP element with the given value
        """
        self.element.text = value
    
    @requires_valid_session
    def select_key(self, key: Any = None):
        """
        Selects a key in the SAP element.
        
        If the element is a table, it selects the first row by default.
        """
        if key:
            self.element.select(key)
        else:
            self.element.select()
            
    @requires_valid_session
    def set_key(self, key: Any) -> None:
        """
        Sets a key in the SAP element.
        
        This is typically used for input fields or selection boxes.
        """
        self.element.key = key
        
    @requires_valid_session
    def set_focus(self) -> None:
        """
        Sets focus on the SAP element.
        
        This is useful for ensuring that the element is ready for input or interaction.
        """
        self.element.setFocus()
        
    def get_tooltip(self) -> str:
        """
        Returns the tooltip of the SAP element.
        
        This is useful for getting additional information about the element.
        """
        return self.element.tooltip
    
    def is_selected(self) -> bool:
        """
        Checks if the SAP element is selected.
        
        Returns:
            bool: True if the element is selected, False otherwise.
        """
        if hasattr(self.element, "Selected"):
            return self.element.Selected
        
        elif hasattr(self.element, "Checked"):
            return self.element.Checked
        
        else:
            raise SapAttributeNotFoundException("Selected/Checked")
    
    @requires_valid_session
    def select(self) -> None:
        """
        Checks the SAP element (checkbox/radio).
        
        This method is used to select a checkbox or radio button in the SAP GUI.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a 'Selected' or 'Checked' attribute.
        """
        if hasattr(self.element, "Selected"):
            self.element.Selected = True
        elif hasattr(self.element, "Checked"):
            self.element.Checked = True
        else:
            raise SapAttributeNotFoundException("Selected/Checked")
    
    @requires_valid_session
    def toggle_select(self) -> None:
        """
        Changes the selection state of the SAP element.
        
        This method toggles the selection state of a checkbox or radio button in the SAP GUI.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a 'Selected' or 'Checked' attribute.
        """
        if hasattr(self.element, "Selected"):
            self.element.Selected = not self.element.Selected
        elif hasattr(self.element, "Checked"):
            self.element.Checked = not self.element.Checked
        else:
            raise SapAttributeNotFoundException("Selected/Checked")

    def get_element_type(self) -> str:
        """
        Returns the type of the SAP element.
        
        This is useful for understanding what kind of element it is (e.g., button, input field, etc.).
        """
        return self.element.Type
    
    @requires_valid_session
    def get_parent(self) -> 'SapElement':
        """
        Returns the parent of the SAP element.
        
        This is useful for navigating the SAP GUI hierarchy.
        
        Returns:
            SapElement: The parent element.
        """
        parent_element = self.element.Parent
        return SapElement(self.session, parent_element)
    
    @requires_valid_session
    def get_children(self) -> list['SapElement']:
        """
        Returns the children of the SAP element.
        
        This is useful for navigating the SAP GUI hierarchy.
        
        Returns:
            list[SapElement]: A list of child elements.
        """
        children = self.element.Children
        return [SapElement(self.session, child) for child in children]
    
    @requires_valid_session
    def get_column(self) -> int:
        """
        Returns the column index of the SAP element.
        
        This is useful for identifying the position of the element in a table or grid.
        
        Returns:
            int: The column index of the element.
        """
        if hasattr(self.element, "Column"):
            return self.element.Column
        
        column_pattern = re.compile(r'[^/]\[(\d+)\s*,\s*\d+]+$')
        column_match = column_pattern.search(self.id)
        
        if column_match:
             column = column_match.group(1)
             return int(column)
        
        raise SapAttributeNotFoundException("Column")
    
    @requires_valid_session
    def get_row(self) -> int:
        """
        Returns the row index of the SAP element.
        
        This is useful for identifying the position of the element in a table or grid.
        
        Returns:
            int: The row index of the element.
        """
        if hasattr(self.element, "Row"):
            return self.element.Row
        
        row_pattern = re.compile(r'[^/]\[\d+\s*,\s*(\d+)]+$')
        row_match = row_pattern.search(self.id)
        
        if row_match:
             row = row_match.group(1)
             return int(row)
        
        raise SapAttributeNotFoundException("Row")
    
    def is_scrollable(self) -> bool:
        """
        Checks if the SAP element is scrollable.
        
        This is useful for determining if the element can be scrolled to view more content.
        
        Returns:
            bool: True if the element is scrollable, False otherwise.
        """
        return hasattr(self.element, 'verticalScrollbar')
    
    def get_scroll_position(self, max_or_min: Optional[Literal['max', 'min']] = None):
        """
        Scrolls the SAP element to the maximum or minimum position.
        
        Args:
            max_or_min (Optional[Literal['max', 'min']]): 
                If 'max', scrolls to the maximum position.
                If 'min', scrolls to the minimum position.
                Defaults to None, which scrolls to the maximum position.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a vertical scrollbar.
        """
        if not self.is_scrollable():
            raise SapAttributeNotFoundException("verticalScrollbar")
        
        scrollbar = self.element.verticalScrollbar
        
        if max_or_min == 'min':
            return scrollbar.minimum
        if max_or_min == 'max':
            return scrollbar.maximum
        else:
            return scrollbar.position
    
    @requires_valid_session
    def scroll_to_position(self, position: int) -> None:
        """
        Scrolls the SAP element to the specified position.
        
        Args:
            position (int): The position to scroll to.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a vertical scrollbar.
        """
        if not self.is_scrollable():
            raise SapAttributeNotFoundException("verticalScrollbar")
        
        scrollbar = self.element.verticalScrollbar
        scrollbar.position = position
    
    @requires_valid_session
    def scroll_to_relative_position(self, direction: Literal['up', 'down'], amount: int = 0) -> bool:
        """
        Scrolls the SAP element vertically by a specified amount in the given direction.
        
        Args:
            direction (Literal['up', 'down']): 
                The direction to scroll. Must be either 'up' or 'down'.
            amount (int): 
                The amount to scroll. Defaults to 0, which means no scrolling.
            
        Returns:
            bool: True if the scroll was successful, False if it would exceed the bounds.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a vertical scrollbar.
            ValueError: If the direction is not 'up' or 'down'.
        """
        if not self.is_scrollable():
            raise SapAttributeNotFoundException("verticalScrollbar")
        
        current_position = self.get_scroll_position()
        max_position = self.get_scroll_position(max_or_min='max')
        min_position = self.get_scroll_position(max_or_min='min')
        
        if direction == 'up':
            new_position = current_position - amount
            
            if new_position < min_position:
                return False
            
        elif direction == 'down':
            new_position = current_position + amount
            
            if new_position > max_position:
                return False
            
        else:
            raise ValueError("Direction must be 'up' or 'down'.")
        
        self.element.verticalScrollbar.position = new_position
        return True
