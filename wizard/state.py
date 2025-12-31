class WizardState:
    def __init__(self):
        self.file_path = None
        self.cscs_sheet = None
        self.ixtrac_sheet = None

        self.headers = []
        self.mapping = {
            "sheet": None,
            "name": None,
            "chn": None,
            "membercode_out": None,
            "status_out": None,
        }
        self.mapping_name = None
