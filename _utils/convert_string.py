class Convert_string:
    def convert_string(self, string):
        if isinstance(string, str):
            string.replace(" ", "")
            int(string)