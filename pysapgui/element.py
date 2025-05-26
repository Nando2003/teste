import re
import win32com.client as client
from typing import Optional, Any, Union, Literal, Sequence, Generator, TYPE_CHECKING

if TYPE_CHECKING:
    from pysapgui.item_element import ItemElement

from pysapgui.session import Session
from pysapgui.utils import search_path, check_element_attribute
from pysapgui.exceptions import SapAttributeNotFoundException


class Element:
    """
    Represents a SAP GUI element.
    
    This class provides methods to interact with SAP GUI elements, such as filling input fields,
    clicking buttons, and navigating through tables and grids.
    It is designed to be used with the SAP GUI Scripting API, allowing automation of SAP GUI tasks.
    
    Attributes:
        session (Session): The SAP session associated with this element.
        element (client.CDispatch): The underlying COM object representing the SAP GUI element.
    """
    @staticmethod
    def each_table_row(
        parent: 'Element',
        column_limit: Optional[int] = None,
        return_empty_rows: bool = False
    ) -> Generator[Sequence['Element'], None, None]:
        """
        Yields each row as a list of GuiTableControl.
        """
        table = parent.element
        columns_count = table.Columns.Count
        
        if column_limit is not None:
            columns_count = min(columns_count, column_limit)

        for row in range(table.Rows.Count):
            row_elements = []
            
            for column in range(columns_count):
                try:
                    cell = table.Columns.elementAt(column).elementAt(row)
                    cell_element = Element(parent.session, cell)
                    try:
                        cell_element.set_column_title(table.Columns.elementAt(column).Title)
                    except Exception:
                        pass
                    row_elements.append(cell_element)
                    
                except Exception:
                    continue

            if not return_empty_rows:
                if all(
                    (getattr(cell, 'get_text', lambda: '')() == '' or getattr(cell, 'get_text', lambda: None)() is None)
                    for cell in row_elements
                ):
                    continue

            if not row_elements:
                break
            
            yield row_elements
        
    def __init__(self, session: Session, element: client.CDispatch):
        self.session = session
        self.element = element
        self.__column_title = None
        
    def __getattr__(self, name: Any) -> Any:
        try:
            return getattr(self.element, name)
        except AttributeError as e:
            raise SapAttributeNotFoundException(name) from e
    
    def __eq__(self, value: object) -> bool:
        return isinstance(value, Element) and (
            self.get_id() == value.get_id() and 
            self.get_type() == value.get_type()
        )
    
    def get_id(self) -> str:
        """
        Returns the ID of the SAP element.
        
        This is useful for identifying the element in the SAP GUI hierarchy.
        
        Returns:
            str: The ID of the SAP element.
        """
        return str(self.element.Id)
    
    def get_type(self) -> str:
        """
        Returns the type of the SAP element.
        
        This is useful for identifying the type of the element in the SAP GUI hierarchy.
        
        Returns:
            str: The type of the SAP element.
        """
        return str(self.element.Type)
    
    @check_element_attribute
    def get_column_title(self) -> str:
        """
        Returns the title of the SAP element.
        
        This is useful for getting the label or title of the element.
        
        Returns:
            str: The title of the SAP element.
        """
        if not self.__title:
            raise SapAttributeNotFoundException("Title")
        
        return str(self.__column_title)
    
    def set_column_title(self, column_title: str) -> None:
        """
        Sets the title of the SAP element.
        
        This is useful for updating the label or title of the element.
        
        Args:
            title (str): The new title for the SAP element.
        """
        self.__column_title = column_title
    
    @check_element_attribute
    def get_text(self) -> str:
        """
        Returns the text of the SAP element.
        
        This is useful for getting the label or title of the element.
        
        Returns:
            str: The text of the SAP element.
        """
        return str(self.element.text)
    
    @check_element_attribute
    def fill(self, value: Any) -> None:
        """
        Fills the SAP element with the given value.
        
        This is typically used for input fields or selection boxes.
        
        Args:
            value (Any): The value to fill in the SAP element.
        """
        self.element.text = value
        
    @check_element_attribute
    def click(self) -> None:
        """
        Simulates a click on the SAP element.
        
        This is useful for buttons or clickable elements in the SAP GUI.
        """
        self.element.press()
    
    @check_element_attribute
    def set_focus(self) -> None:
        """
        Sets focus on the SAP element.
        
        This is useful for ensuring that the element is ready for input or interaction.
        """
        self.element.setFocus()
        
    @check_element_attribute
    def select_key(self, key: Any = None) -> None:
        """
        Selects a key in the SAP element.
        
        If the element is a table, it selects the first row by default.
        
        Args:
            key (Any): The key to select in the SAP element. Defaults to None.
        """
        self.element.select(key) if key else self.element.select()

    @check_element_attribute
    def set_key(self, key: Any) -> None:
        """
        Sets a key in the SAP element.
        
        This is typically used for input fields or selection boxes.
        """
        self.element.key = key
    
    @check_element_attribute
    def get_tooltip(self) -> str:
        """
        Returns the tooltip of the SAP element.
        
        This is useful for getting additional information about the element.
        """
        return self.element.tooltip
    
    @check_element_attribute     
    def is_selected(self) -> bool:
        """
        Checks if the SAP element (checkbox/radio) is selected.
        
        Returns:
            bool: True if the element is selected, False otherwise.
        """
        if hasattr(self.element, "Selected"):
            return self.element.Selected
        elif hasattr(self.element, "Checked"):
            return self.element.Checked
        raise SapAttributeNotFoundException("Selected/Checked")
    
    @check_element_attribute
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
        raise SapAttributeNotFoundException("Selected/Checked")
    
    @check_element_attribute
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
        raise SapAttributeNotFoundException("Selected/Checked")

    @check_element_attribute
    def get_parent(self) -> 'Element':
        """
        Returns the parent of the SAP element.
        
        This is useful for navigating the SAP GUI hierarchy.
        
        Returns:
            Element: The parent element wrapped in an Element instance.
        """
        parent_element = self.element.Parent
        return Element(self.session, parent_element)
    
    @check_element_attribute
    def get_children(self) -> list['Element']:
        """
        Returns the children of the SAP element.
        
        This is useful for navigating the SAP GUI hierarchy.
        
        Returns:
            list[Element]: A list of child elements wrapped in Element instances.
        """
        children = self.element.Children
        return [Element(self.session, child) for child in children]
    
    @check_element_attribute
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
    
    @check_element_attribute
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
    
    @check_element_attribute
    def is_scrollable(self) -> bool:
        """
        Checks if the SAP element is scrollable.
        
        This is useful for determining if the element can be scrolled to view more content.
        
        Returns:
            bool: True if the element is scrollable, False otherwise.
        """
        return hasattr(self.element, 'verticalScrollbar')
    
    @check_element_attribute
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
    
    @check_element_attribute
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
    
    @check_element_attribute
    def scroll_to_absolute_position(self, position: int) -> None:
        """
        Scrolls the SAP element to an absolute position.
        
        Args:
            position (int): The absolute position to scroll to.
        
        Returns:
            bool: True if the scroll was successful, False if it would exceed the bounds.
        
        Raises:
            SapAttributeNotFoundException: If the element does not have a vertical scrollbar.
        """
        if not self.is_scrollable():
            raise SapAttributeNotFoundException("verticalScrollbar")
        
        self.element.verticalScrollbar.position = position
    
    def find_partial_element(self, re_element_id: str) -> Optional['Element']:
        """
        Finds a child element using a partial (regex) path.

        Args:
            re_element_id (str): The partial or regex path of the child element.

        Returns:
            Optional[Element]: The first matching child element wrapped in an Element instance, or None if not found.
        """
        results = search_path(
            self.element,
            re_path=re_element_id,
            return_all=False,
            element_wrapper=lambda elem: Element(self.session, elem)
        )
        return results if isinstance(results, Element) else None
          
    def each_row(
        self, 
        column_limit: Optional[int] = None, 
        return_empty_rows: bool = True
    ) -> Generator[Union[Sequence['Element'], Sequence[ItemElement]], None, None]:
        """
        Iterates over each row of the current SAP element, yielding either ItemElement
        or Element instances, depending on the underlying SAP control type.

        This method provides a uniform interface to iterate over rows for different
        table-like SAP GUI controls, adapting the behavior as follows:

        - For SAP GridView controls ("GridViewCtrl" in element text):
            Uses `GridViewElement.each_row`, yielding lists of GridViewElement for each row.
            Each item represents a cell in the grid.
        
        - For SAP TableTree controls ("TableTreeCtrl" in element text):
            Uses `TableTreeElement.each_row`, yielding lists of TableTreeElement for each row.
            Each item represents a node or cell in the table tree.
        
        - For classic SAP Table controls ("GuiTableControl" in element type):
            Uses `Element.each_table_row`, yielding lists of Element, one per cell, for each row.
            Handles both the grid and plain table visualizations.
        
        - For other SAP elements that expose row/column information:
            Falls back to an internal row generator (`__rows_generator`), which yields
            lists of Element for each detected row.
        
        Args:
            column_limit (Optional[int]): 
                The maximum number of columns to return for each row. 
                If None, returns all columns.
            return_empty_rows (bool): 
                If True, includes empty rows in the iteration.
                If False, skips rows where all cells are empty or blank.
        
        Yields:
            Sequence[Element] or Sequence[ItemElement]: 
                For each row, yields a sequence (usually a list) of Element or ItemElement 
                instances representing the cells of that row.
        """
        from pysapgui.item_element import GridViewElement, TableTreeElement
        
        element_text = self.get_text()
        element_type = self.get_type()
        
        if 'GridViewCtrl' in element_text:
            yield from GridViewElement.each_row(
                parent=self, 
                column_limit=column_limit, 
                return_empty_rows=return_empty_rows
            )
            
        elif 'TableTreeCtrl' in element_text:
            yield from TableTreeElement.each_row(
                parent=self, 
                column_limit=column_limit, 
                return_empty_rows=return_empty_rows
            )
        
        elif 'GuiTableControl' in element_type:
            yield from Element.each_table_row(
                parent=self, 
                column_limit=column_limit, 
                return_empty_rows=return_empty_rows
            )
        
        yield from self.__rows_generator(
            column_limit=column_limit, 
            remove_empty_rows=not return_empty_rows
        )
        
    def __rows_generator(self, column_limit: Optional[int] = None, remove_empty_rows: bool = True):
        rows = {}
        
        for element in self.get_children():
            try:
                r = element.row
                c = element.column
            except Exception:
                continue
            
            if column_limit is not None and c >= column_limit:
                continue
            
            if r not in rows:
                rows[r] = []
            
            if not getattr(element, "text", None):
                continue
            
            rows[r].append(element)
            
        for r in sorted(rows.keys()):
            row_elems = sorted(rows[r], key=lambda e: e.column)
            
            if column_limit is not None:
                row_elems = row_elems[:column_limit]
                
            if remove_empty_rows and not row_elems:
                continue
            
            yield row_elems