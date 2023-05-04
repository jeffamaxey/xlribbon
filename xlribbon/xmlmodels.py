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
        strout = f"<{name} "
        for key, field in self.dict(exclude={"children"}, exclude_none=True).items():
            strout += f"{key}='{field}' "
        if self.children:
            strout += ">"
            for child in self.children:
                strout += child.xml()
            strout += f"</{name}>"
        else:
            strout += "/>"
        return strout

    def __excel_name__(self):
        name = self.__repr_name__()
        return name[:1].lower() + name[1:]

    def get_callbacks(self):
        """Get all callbacks from the model defintions."""
        callbacks = list(self.dict(include=CALLBACKS, exclude_none=True).values())
        if self.children:
            [callbacks.extend(child.get_callbacks()) for child in self.children]
        return callbacks

    def get_getters(self, module_name: str = "xlribbon"):
        """Look for unset getters and set to a default value."""
        auto_getters = {}
        getters = self.dict(include=GETTERS)
        for key, value in getters.items():
            if value is None:
                auto_getters[f"{module_name}.{self.id}_{key}"] = key
                self.__setattr__(key, f"{module_name}.{self.id}_{key}")

        if self.children:
            [auto_getters.update(child.get_getters()) for child in self.children]
        return auto_getters

    def get_setters(self, module_name: str = "xlribbon"):
        """Look for unset setters and set to a default value."""
        auto_setters = {}
        setters = self.dict(include=SETTERS)
        for key, value in setters.items():
            if value is None:
                auto_setters[f"{module_name}.{self.id}_{key}"] = key
                self.__setattr__(key, f"{module_name}.{self.id}_{key}")

        if self.children:
            [auto_setters.update(child.get_setters()) for child in self.children]
        return auto_setters

    def get_images(self):
        """Look for images."""
        setters = self.dict(include={"image"})
        images = [value for key, value in setters.items() if value]
        if self.children:
            [images.extend(child.get_images()) for child in self.children]
        return images


class UI(XMLBase):
    enabled: Optional[str] = None
    label: str
    image: Optional[str] = None
    imageMso: Optional[str] = None
    size: str
    onAction: str
    screentip: Optional[str] = None
    supertip: Optional[str] = None


class CheckBox(UI):
    getPressed: str
    onAction: Optional[str] = None


class Button(UI):
    onAction: str


class Item(XMLBase):
    label: str


class ComboBox(UI):
    children: List[Item]
    sizeString: str
    onChange: str


class EditBox(XMLBase):
    label: str
    onChange: str


class ButtonGroup(XMLBase):
    children: List[Button]


class LabelControl(XMLBase):
    getText: Optional[str] = None
    setText: Optional[str] = None


class Group(XMLBase):
    label: str
    children: List[Union[Button, ButtonGroup, CheckBox, ComboBox, EditBox, LabelControl]]


class Tab(XMLBase):
    label: str
    children: List[Group]


class Tabs(XMLBase):
    label: str
    children: List[Tab]
