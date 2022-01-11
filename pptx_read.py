# Import external modules.
from pptx import Presentation


class PPTXRead:

    def __init__(self, in_file_name):
        self.presentation = Presentation(in_file_name)
        return

    def get_table_values(self, in_slide_nr) -> list:
        out_list = []
        a_slide = self.presentation.slides[in_slide_nr]
        for a_shape in a_slide.shapes:
            c = 1
        a_table = a_slide.shapes[1].table
        for row in a_table.rows:
            for column in row.cells:
                print(column.text_frame.text)
        return out_list


if __name__ == '__main__':
    # Test class.
    a_pptx = PPTXRead('data_ppt.pptx')
    a_pptx.get_table_values(1)
