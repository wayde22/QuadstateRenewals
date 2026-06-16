class MissingDependencyError(RuntimeError):
    def __init__(self, package_name, install_hint, missing_module=None):
        self.package_name = package_name
        self.install_hint = install_hint
        self.missing_module = missing_module or package_name
        super().__init__(
            f"Missing dependency '{self.missing_module}'. "
            f"Install or package '{self.install_hint}'."
        )
