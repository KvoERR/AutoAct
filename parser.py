class Parser:
    def __init__(self, data_path, template_path, 
                template_list, field_list, word_list, 
                base_template, modified_template, modification):
            self.data_dir = data_path
            self.templates_dir = template_path

            self.template_list = template_list
            self.field_list = field_list
            self.word_packs = word_list

            self.base_template = base_template
            self.modified_template = modified_template
            self.modification = modification

            print(self.__dict__)
