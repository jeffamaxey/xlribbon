from datetime import datetime

from .templates import TEMPLATES


class ConfigurationError(Exception):
    pass


class Ribbon:
    def __init__(self, model, router) -> None:
        self.model = model
        self.router = router
        self._check_model_and_router()
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

    def xml(self):
        """Generate the model ui xml and required framing."""
        return (
            f'<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">\n\t'
            f'<ribbon startFromScratch="false">{self.model.xml()}'
            f"</ribbon>\n</customUI>"
        )

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

    def write_routes(self):
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
