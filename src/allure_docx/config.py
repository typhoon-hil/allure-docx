import os
from configparser import ConfigParser
from enum import Enum
from enum import EnumMeta

_config_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), "config")


class ConfigTagsEnumMeta(EnumMeta):
    """
    MetaClass to make "in" operator possible with strings for ConfigTags enum.
    """

    def __contains__(cls, item):
        return isinstance(item, cls) or item in list(cls.__members__.keys())


class ConfigTags(Enum, metaclass=ConfigTagsEnumMeta):
    """
    Configuration tags, that can be used to create a ReportConfig. Holds the ini file paths for each config tag.
    """

    STANDARD = os.path.join(_config_dir, "standard.ini")
    STANDARD_ON_FAIL = os.path.join(_config_dir, "standard_on_fail.ini")
    COMPACT = os.path.join(_config_dir, "compact.ini")
    NO_TRACE = os.path.join(_config_dir, "no_trace.ini")

    @staticmethod
    def get_names() -> [str]:
        """
        Returns all tags in lower case as string array.
        """
        return [v.name.lower() for v in ConfigTags]


class ReportConfig(dict):
    """
    A report config extending a dictionary for simple .ini file importing in the correct format.
    """

    def __init__(self, tag: ConfigTags = None, config_file: str = None):
        """
        Create a ReportConfig from either a tag defined in "ConfigTags" enum or from a path to a .ini configuration file

        Parameters:
            tag : Tag defined in ConfigTags.
            config_file : Path to a .ini configuration file.
        """

        super().__init__()
        if tag and config_file:
            raise ValueError("Cannot initialize ReportConfig with both tag and file.")

        self.config_parser = ConfigParser()
        self.config_parser.optionxform = str
        if tag:
            self.config_parser.read(tag.value)
        else:
            standard_file = ConfigTags.STANDARD.value
            self.config_parser.read(standard_file)
            if config_file:
                if config_file is not standard_file:
                    self.config_parser.read(config_file)
        self._build_dict()

    def _build_dict(self):
        """
        Creates the dictionary from the config_parser. Parameter in "info" and "labels" are parsed to each section
        (failed, broken, passed, skipped, unknown)
        """

        def transform_by_status_to_dict(section):
            section_old = self[section]
            self[section] = {}
            self[section]["failed"] = []
            self[section]["broken"] = []
            self[section]["passed"] = []
            self[section]["skipped"] = []
            self[section]["unknown"] = []
            for key in section_old:
                if "f" in section_old[key]:
                    self[section]["failed"].append(key)
                if "b" in section_old[key]:
                    self[section]["broken"].append(key)
                if "p" in section_old[key]:
                    self[section]["passed"].append(key)
                if "s" in section_old[key]:
                    self[section]["skipped"].append(key)
                if "u" in section_old[key]:
                    self[section]["unknown"].append(key)

        self.update({s: dict(self.config_parser.items(s)) for s in self.config_parser.sections()})
        transform_by_status_to_dict("info")
        transform_by_status_to_dict("labels")
