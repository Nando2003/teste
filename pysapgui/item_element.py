from abc import ABC, abstractmethod
from pysapgui.element import Element
from typing import Any, Generator, Sequence, Optional

from pysapgui.exceptions import (
    TableTreeSelectAllNotSupportedException,
    TableTreeColumnSelectionException
)


class ItemElement(ABC):
    """
    Represents a generic item element within a SAP GUI structure.

    This class is intended to serve as a base or interface for more specialized
    item elements (such as table cells, grid view items, or tree nodes) in the SAP GUI.
    It can be extended to encapsulate additional behaviors and attributes relevant to
    specific item types within complex SAP GUI components.

    Typically, an ItemElement will reference a parent Element and its corresponding
    position (row, column, or key) within the parent control.

    Attributes:
        (To be defined in subclasses according to the specific SAP GUI control type.)
    """ 
    
    @staticmethod
    @abstractmethod
    def each_row(parent: Element, column_limit: Optional[int] = None, return_empty_rows: bool = True) -> Generator[Sequence['ItemElement'], None, None]:
        ...
    
    def __init__(self, parent: Element, row: Any, col: Any) -> None:
        self.row = row
        self.col = col
        self.parent = parent
    
    @abstractmethod
    def get_header(self) -> str:
        """
        Get the header or label of the item element.
        
        Returns:
            str: The header or label of the item element.
        """
        pass
    
    @abstractmethod
    def get_text(self) -> str:
        """
        Get the text content of the item element.
        
        Returns:
            str: The text content of the item element.
        """
        pass
    
    @abstractmethod
    def select(self) -> None:
        """
        Select the item element.
        
        This method should implement the logic to select the item element
        within its parent control.
        """
        pass
    
    @abstractmethod
    def select_column(self) -> None:
        """
        Select the column of the item element.
        
        This method should implement the logic to select the column
        of the item element within its parent control.
        """
        pass
    
    @abstractmethod
    def select_row(self) -> None:
        """
        Select the row of the item element.
        
        This method should implement the logic to select the row
        of the item element within its parent control.
        """
        pass
    
    @abstractmethod
    def select_all(self) -> None:
        """
        Select all items in the item element.
        
        This method should implement the logic to select all items
        within the item element's parent control.
        """
        pass
    
    @abstractmethod
    def double_click(self) -> None:
        """
        Double-click the item element.
        
        This method should implement the logic to double-click the item element
        within its parent control.
        """
        pass
    
    @abstractmethod
    def clear_selection(self) -> None:
        """
        Clear the selection of the item element.
        
        This method should implement the logic to clear the selection
        of the item element within its parent control.
        """
        pass


class GridViewElement(ItemElement):
    """
    Represents a grid view item element in the SAP GUI.

    This class extends ItemElement to provide specific behaviors and attributes
    for grid view items, such as selecting rows, columns, and handling text content.
    """
        
    @staticmethod
    def each_row(parent: Element, column_limit: Optional[int] = None, return_empty_rows: bool = True) -> Generator[Sequence['GridViewElement'], None, None]:
        """
        Yields each row as a list of GridViewElement.
        """
        element = parent.element
        columns_count = element.ColumnCount if column_limit is None else column_limit
            
        for row in range(element.RowCount):
            row_elements = [
                GridViewElement(parent, row, element.ColumnOrder.Item(col))
                for col in range(columns_count)
            ]
            
            if not return_empty_rows:
                if all((cell.get_text() is None or str(cell.get_text()).strip() == "") for cell in row_elements):
                    continue
                
            yield row_elements
    
    def get_header(self) -> str:
        return self.parent.GetDisplayedColumnTitle(self.col)
    
    def get_text(self) -> str:
        return self.parent.GetCellValue(self.row, self.col)
    
    def select(self) -> None:
        self.parent.SetCurrentCell(self.row, self.col)
    
    def select_column(self) -> None:
        self.parent.SelectColumn(self.col)
    
    def select_row(self) -> None:
        self.parent.selectedRows = self.row # type: ignore[assignment]
    
    def select_all(self) -> None:
        self.parent.SelectAll()
    
    def double_click(self) -> None:
        self.parent.DoubleClickCell(self.row, self.col)
    
    def clear_selection(self) -> None:
        self.parent.ClearSelection()
        

class TableTreeElement(ItemElement):
    """
    Represents a table tree item element in the SAP GUI.

    This class extends ItemElement to provide specific behaviors and attributes
    for table tree items, such as selecting rows, columns, and handling text content.
    """
    
    @staticmethod
    def each_row(parent: Element, column_limit: Optional[int] = None, return_empty_rows: bool = True) -> Generator[Sequence[ItemElement], None, None]:
        """
        Yields each row as a list of TableTreeElement.
        """
        element = parent.element
        column_names = list(element.GetColumnNames())
        
        if column_limit is not None:
            column_names = column_names[:column_limit]

        for row_key in element.GetAllNodeKeys():
            row_elements = [
                TableTreeElement(parent, row_key, col_name)
                for col_name in column_names
            ]

            if not return_empty_rows:
                
                if all((cell.get_text() is None or str(cell.get_text()).strip() == "") for cell in row_elements):
                    continue
            
            yield row_elements
    
    def get_header(self) -> str:
        return self.parent.GetColumnTitleFromName(self.col)
    
    def get_text(self) -> str:
        return self.parent.GetItemText(self.row, self.col)
    
    def select(self) -> None:
        self.parent.SelectItem(self.row, self.col)
    
    def select_column(self) -> None:
        try:
            self.parent.SelectColumn(self.col)
            
        except Exception as e:
            raise TableTreeColumnSelectionException(self.col) from e
    
    def select_row(self) -> None:
        self.parent.SelectNode(self.row)
    
    def select_all(self) -> None:
        raise TableTreeSelectAllNotSupportedException
    
    def double_click(self) -> None:
        self.parent.DoubleClickItem(self.row, self.col)
    
    def clear_selection(self) -> None:
        self.parent.unSelectAll()