class WizardState:
    def __init__(self):
        self.file_path = None

        # runtime sheet selections
        self.cscs_sheet = None
        self.ixtrac_sheet = None

        # discovered headers from IX TRAC sheet
        self.headers = []

        # mapping to be saved
        self.mapping = {
            "cscs_sheet": None,
            "ixtrac_sheet": None,
            "name": None,
            "chn": None,
            "membercode_out": None,
            "status_out": None,
        }
