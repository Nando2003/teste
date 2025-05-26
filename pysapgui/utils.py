import re
from pysapgui.exceptions import SapAttributeNotFoundException
from typing import Any, Union, Optional, Callable, TYPE_CHECKING

if TYPE_CHECKING:
    from win32com.client import CDispatch
    
    
def search_path(
    base_element: 'CDispatch', 
    re_path: str,
    return_all: bool = False,
    element_wrapper: Any = None
) -> list:
    parts = re_path.split('/')
    current_level = [base_element]
    results = []
    
    for i, part in enumerate(parts):
        next_level = []
        pattern = re.sub(r'(?<!\\)(\[)', r'\\[', part)
        
        for elem in current_level:
            found = search_element(elem, pattern, (i == len(parts) - 1) and return_all)
            
            if found:
                if isinstance(found, list):
                    next_level.extend(found)
                else:
                    next_level.append(found)
        
        if not next_level:
            return []
        
        current_level = next_level
        
    for elem in current_level:
        results.append(element_wrapper(elem) if element_wrapper else elem)
        
    if not return_all and results:
        return results[0]
    
    return results
        

def search_element(
    element: 'CDispatch',
    pattern: str,
    return_all: bool,
) -> Optional[Union['CDispatch', list['CDispatch']]]:
    found_elements = []
    
    if hasattr(element, 'Children'):
        try:
            children_iter = iter(element.Children)
            
        except TypeError:
            return None
    
        for child in children_iter:
            if re.search(pattern, child.Id, re.IGNORECASE):
                if not return_all:
                    return child
                
                found_elements.append(child)
                    
            sub_found = search_element(child, pattern, return_all)
            
            if sub_found:
                if isinstance(sub_found, list):
                    if not return_all:
                        return sub_found[0]
                    
                    found_elements.extend(sub_found)

                else:
                    if not return_all:
                        return sub_found
                    
                    found_elements.append(sub_found)
                
    if return_all and found_elements:
        return found_elements
    
    return None if not return_all else []


def check_element_attribute(method: Callable):
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        
        except AttributeError as e:

            if isinstance(e, SapAttributeNotFoundException):
                raise
            
            msg = e.args[0] if e.args else str(e)
            match = re.search(r"has no attribute '([^']+)'", msg)
            
            if match:
                attribute = match.group(1)
            else:
                attribute = msg
                
            raise SapAttributeNotFoundException(attribute) from e
        
    return wrapper
