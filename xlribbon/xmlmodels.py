from typing import Any, List, Optional, Union

from pydantic import BaseModel

CALLBACKS = {"onAction", "getPressed"}
GETTERS = {"getVisible", "getText", "getLabel"}
SETTERS = {"onChange", "setText"}


class XMLBase(BaseModel):
    id: str
    children: Union[None, List[Any]] = None

    def xml(self):
        name = self.__excel_name__()
        strout = f"\n<{name} \n"
        for key, field in self.dict(exclude={"children"}).items():
            strout += f"\t {key}='{field}' \n"
        strout += ">"
        if self.children:
            for child in self.children:
                strout += child.xml()
            strout += f"\n</{name}>"
        return strout

    def __excel_name__(self):
        name = self.__repr_name__()
        return name[:1].lower() + name[1:]

    def get_callbacks(self):
        """Get all callbacks from the model defintions."""
        callbacks = list(self.dict(include=CALLBACKS).values())
        if self.children:
            [callbacks.extend(child.get_callbacks()) for child in self.children]
        return callbacks

    def get_getters(self):
        """Look for unset getters and set to a default value."""
        auto_getters = {}
        getters = self.dict(include=GETTERS)
        for key, value in getters.items():
            if value is None:
                auto_getters.update({f"{self.id}_{key}": key})
                self.__setattr__(key, f"{self.id}_{key}")

        if self.children:
            [auto_getters.update(child.get_getters()) for child in self.children]
        return auto_getters

    def get_setters(self):
        """Look for unset setters and set to a default value."""
        auto_setters = {}
        setters = self.dict(include=SETTERS)
        for key, value in setters.items():
            if value is None:
                auto_setters.update({f"{self.id}_{key}": key})
                self.__setattr__(key, f"{self.id}_{key}")

        if self.children:
            [auto_setters.update(child.get_setters()) for child in self.children]
        return auto_setters


class UI(XMLBase):
    label: str
    imageMso: str
    size: str
    onAction: str
    screentip: Optional[str] = None
    supertip: Optional[str] = None


class Tab(XMLBase):
    label: str
    children: List[Any]


class Tabs(XMLBase):
    label: str
    children: List[Tab]


class CheckBox(UI):
    getPressed: str


class Button(UI):
    getPressed: str


class LabelControl(XMLBase):
    getText: Optional[str] = None
    setText: Optional[str] = None
