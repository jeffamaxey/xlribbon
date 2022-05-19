# xlribbon

xlribbon is a lightweight package for generating a custom excel ribbon from python code. It accompanies xlwings and enhanced the custom addin capabilities for larger projects, lending from ideas used in FASTApi or JS frontend frameworks, and removing the need to ever manually adjusting your ribbon.

## User Interface
 The package leverages pydantic for defining the python representation of the excel ribbon objects and automatically outputs the xml version.

```python
from xlribbon.xmlmodels import Tabs, Tabs, Button
custom_ui = Tabs(
    id="a",
    label="Maintab",
    children=[
        Tab(
            id="1",
            label="tab_label",
            children=[
                Group(
                    id="maingroup",
                    label="MainGroup",
                    children=[
                        Button(
                            id="mycheckbox",
                            label="Click this",
                            onAction="some_callback",
                            size="large",
                            image="Image1",
                        ),
                        LabelControl(id="labelcontrol")
                )
            ],
        )
    ],
)

custom_ui.xml()
```

Pydantic takes care of the required attributes and not yet implemented excel objects can be easily created by inheriting from 'XMLBase' or the more elaborate 'UI' class. The code for the Button class is e.g
```python
class Button(UI):
    onAction: str
```

## Router

Besides the xml model, xlribbon provides a Router class which collects python callbacks simple as e.g. FASTApi:
```python
from xlribbon import Router

side_router = Router(prefix="side")

@router.add_route()
def some_callback():
    print("hello")
```

After running the above code, the router has the path "side\some_callback". This is particularly useful if the excel app used several modules or packages as multiple routers can be combined by a simple add_router call:
```python
# Other file
from xlribbon import Router

main_router = Router(prefix="main")
main_router.add_router(side_router)

@router.add_route()
def another_callback():
    print("hello from side")
```
From the combined 'main_router', we can easily generate the required vba code by issuing 'main_router.vba()'.

## Ribbon
Both, the UI and the router are inputs for the 'Ribbon' class, which checks the conformity between router and ui by checking for all required routes. Moreover the ribbon automatically sets "getters" and "setters" for the excel ribbon inputs, which under the hood use the xlings provided .conf file. 

 A ribbon instance allows to generate the custom .xlam file, completely with UI and xlwings RunPython callbacks.

```python
from xlribbon import Ribbon

ribbon = Ribbon(model=custom_ui, router=main_router)
ribbon.make_addin(name="my_custom_addin")
```
After running the above code, the created adding "my_custom_addin.xlam" is available in the working directory.
