class Router:
    def __init__(self, prefix):
        self.views = []
        self.prefix = prefix

    def add_route(self, path, func):
        self.views.append({path:f"{self.prefix}/path", func:func})
