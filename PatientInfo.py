class QueryInfo:
    def __init__(self, first_name, last_name, medicare_number, dob, from_dos, to_dos, invalid):
        self.medicare_number = medicare_number
        self.first_name = first_name
        self.last_name = last_name
        self.dob = dob
        self.from_dos = from_dos
        self.to_dos = to_dos
        self.invalid = invalid