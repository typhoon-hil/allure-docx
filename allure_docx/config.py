import os
from configparser import ConfigParser
import sys


class ReportConfig(dict):
    """
    A report config extending a dictionary for simple .ini file importing in the correct format.
    """

    def read_config_from_file(self, path):
        """
        Read the report config from a file.
        Parameters:
            path : path to a configuration that overwrites the standard configuration

        >>> config = ReportConfig()
        >>> config.read_config_from_file("config/standard.ini", "config/no_trace.ini")
        >>> "description" in config["info"]["failed"]
        True
        >>> "trace" in config["info"]["failed"]
        False
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

        standard_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "config", "standard.ini")
        config_parser = ConfigParser()
        config_parser.read(standard_path)
        if path is not standard_path:
            config_parser.read(path)
        self.update({s: dict(config_parser.items(s)) for s in config_parser.sections()})
        transform_by_status_to_dict("info")
        transform_by_status_to_dict("labels")

