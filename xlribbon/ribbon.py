import logging
import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Union
from xml.dom.minidom import parseString
from zipfile import ZIP_DEFLATED, ZipFile

import xlwings as xw

from .templates import TEMPLATES

logger = logging.getLogger(__name__)


class ConfigurationError(Exception):
    pass


class Ribbon:
    def __init__(self, model, router, img_source: Path = Path("img")) -> None:
        self.model = model
        self.router = router
        self.img_source = img_source
        self._check_model_and_router()
        self._check_required_images()
        self.getters = []
        self.setters = []
        self.model_dict = model.dict()

    def _check_model_and_router(self):
        """Check if all model functions exist in the router."""

        required_callbacks = set(self.model.get_callbacks())
        available_callbacks = set(list(self.router.keys()))
        unavaible_callbacks = required_callbacks - available_callbacks
        if len(unavaible_callbacks) > 0:
            raise ConfigurationError(
                f"Model and Router are not compatible, functions {' ,'.join(list(unavaible_callbacks))} are missing."
            )
        logger.debug("UI model and function router matched.")

    def _check_required_images(self):
        """Check if all required images are available in the 'img' folder."""

        required_images = self.model.get_images()
        images_path = Path().absolute().joinpath(self.img_source)
        for image in required_images:
            if not images_path.joinpath(f"{image}.png").exists():
                raise ConfigurationError(f"Image {image}.png not found in the '{self.img_source}' folder.")
        logger.debug(f"All required images found in '{self.img_source}'.")

    def xml(self):
        """Generate the model ui xml and required framing."""
        xml_string = (
            f'<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">'
            f'<ribbon startFromScratch="false">{self.model.xml()}'
            f"</ribbon></customUI>"
        )

        return parseString(xml_string).toprettyxml()

    def make_getters(self):
        """Build the getter macros."""
        all_getters = self.model.get_getters()
        for key, value in all_getters.items():
            function_body = TEMPLATES[value].format(name=key, function="dosome")
            self.getters.append(f"'AUTOMATIC GETTER for '{key}' \n{function_body}\n\n")

    def make_setters(self):
        """Build the setter macros."""
        all_setters = self.model.get_setters()
        for key, value in all_setters.items():
            function_body = TEMPLATES[value].format(name=key, function="dosome")
            self.setters.append(f"'AUTOMATIC SETTER for '{key}' \n{function_body}\n\n")

    def write_ui(self):
        with open("customUI.xml", mode="x") as f:
            f.write(self.xml())

    def build_routes(self):
        """Generate the required macros."""
        # General functions, ribbon_onload and invalidation
        initializers = f"'xlribbon generated at {datetime.now()}\n\n"
        # Getters and Setters
        getters = "\n".join(self.getters)
        setters = "\n".join(self.setters)
        # Router functions
        routes = "\n\n Automatic routes\n"
        # router.xml()
        return initializers + getters + setters + routes

    def vba_routes(self):
        """Pretty print the vba routes.

        If you cannot access the vba project model, copy the output
        of this function manually to a vba module named 'xlribbon'.
        """
        print(self.build_routes())

    def build_rels(self):
        """Write the required relations for the required images."""
        images = self.model.get_images()
        rel_inital = (
            '<Relationship Id="{image}" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"'
            ' Target="customImages/{image}.png"/>'
        )
        return "\n".join([rel_inital.format(image=image) for image in images])

    def make_addin(self, name: str, update: bool = False):
        """Write all files and make a new package."""
        # Generate Output
        build_dir_name = "./build"
        os.mkdir(build_dir_name)
        logger.debug("Build directory created.")
        # with open(f"{build_dir_name}/routes.txt", mode="x") as f:
        #     f.write(self.build_routes())

        # Prepare the .xlam file
        if update:
            source_file = Path(build_dir_name).joinpath(f"{name}.xlam")
        else:
            source_file = Path(xw.__file__).parent.joinpath("quickstart_addin_ribbon.xlam")
        temp_file = Path(build_dir_name).joinpath(f"{name}_.xlam")
        target_file = Path(build_dir_name).joinpath(f"{name}.xlam")
        shutil.copyfile(source_file, temp_file)
        shutil.copyfile(source_file, target_file)
        logger.debug("Initial .xlam file copied.")
        # Inject VBA Code
        try:
            book = xw.Book(target_file)
            xlribbonmodule = book.api.VBProject.VBComponents("xlribbon")
            xlribbonmodule.CodeModule.AddFromString(self.build_routes())
            logger.debug("VBA code injected into 'xlribbon' module.")
        except Exception:
            logger.error(
                "Injecting VBA code failed. If you want to add the route code manually "
                "call the 'build_routes' method of your ribbon object."
            )

        # Replace .rels and .customUI in the .xlam
        do_not_copy_files = ["customUI/customUI.xml", "customUI/customUI.xml.rels"]

        with ZipFile(temp_file) as source_zip, ZipFile(target_file, "w") as source_zip:
            # Iterate the input files
            logger.debug("Copying standard files.")
            for inzipinfo in source_zip.infolist():
                # Read input file
                with source_zip.open(inzipinfo) as infile:
                    if inzipinfo.filename not in do_not_copy_files:
                        source_zip.writestr(inzipinfo.filename, infile.read(), compress_type=ZIP_DEFLATED)

            logger.debug("Updating UI and relations.")
            source_zip.writestr("customUI/customUI.xml", self.xml(), compress_type=ZIP_DEFLATED)
            source_zip.writestr("customUI/_rels/customUI.xml.rels", self.build_rels(), compress_type=ZIP_DEFLATED)

            # Package all required images
            logger.debug("Updating images.")
            images = self.model.get_images()
            for image in images:
                source_zip.write(
                    filename=Path(self.img_source).joinpath(f"{image}.png"),
                    arcname=f"customUI/images/{image}.png",
                    compress_type=ZIP_DEFLATED,
                )
        logger.debug(".xlam file updated.")

        # Cleanup
        shutil.copyfile(target_file, target_file.parent.parent.joinpath(target_file.name))
        temp_file.unlink()
        target_file.unlink()
        Path(build_dir_name).rmdir()
        logger.debug("Build order cleaned.")
