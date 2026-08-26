from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

def set_vertical_text(cell):
    """Устанавливает вертикальный текст снизу вверх."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), 'btLr')  # btLr — снизу вверх
    tcPr.append(textDirection)