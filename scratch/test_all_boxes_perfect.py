import sys
import os
sys.path.insert(0, r'c:\Users\ACER\Desktop\RANDOM PROJECT\JSON TO DOCX CONVERTER')
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
doc = word.Documents.Add()
sel = word.Selection

font_size = 12.0
font_name = 'Times New Roman'

boxes_test = [
    ('Box 1 (2 rows)', 'native | statue | capital | landscapes | rich | royal | palace | sightseeing'),
    ('Box 2 (2 rows)', 'coastline | eco-friendly | flying cars | harmful | high-speed | kilt | renewable | tattoos'),
    ('Box 3 (2 rows)', 'active | chapped | dim | dirty | fit | healthy | tidy | tired'),
    ('Box 4 (2 rows)', 'coloured vegetables | eye drops | health | healthy diet | lip balm | pimples | soft drinks | sunburn | suncream | vitamins'),
    ('Box 5 (4 rows, 20 items)', 'allergy | exercising | playing sports | runny nose | stomachache | boating | fever | red skin | sleeping | swimming | cough | flu | relaxing | sneezing | walking | cycling | headache | running | sore throat | watching TV'),
    ('Box 6 (6 sentences)', "Because they like doing something useful and helping others. | Yes, it makes a better life and improves the society. | Because volunteering teaches me a lot. | Yes, I've been a volunteer teacher for Street Child Organization. | It helps you stay healthy, increases self-confidence, and makes you happy. | We can donate money or clothes via charitable organisations.")
]

from renderers.exercises.boxes import BoxesRendererMixin
from renderers.common.units_and_colors import cm_to_pt
from renderers.common.typography import TypographyMixin

d = Dummy = type('Dummy', (BoxesRendererMixin, TypographyMixin), {'font_name': font_name, 'font_size': font_size})()
printable_width_pt = cm_to_pt(17.0)

for name, raw in boxes_test:
    words = [w.strip() for w in raw.split('|') if w.strip()]
    cols, grid, box_w_pt, tabs = d.optimize_word_bank_layout(None, words, printable_width_pt)
    num_rows = len(grid)
    
    # Accurate Word TextFrame height formula
    box_h_pt = (num_rows * (font_size * 1.25)) + (font_size * 0.40) + 2.0
    
    # Create Shape in Word
    shape = doc.Shapes.AddShape(5, 0, 0, box_w_pt, box_h_pt)
    tf = shape.TextFrame
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tf.MarginLeft = cm_to_pt(0.20)
    tf.MarginRight = cm_to_pt(0.20)
    tf.WordWrap = 0
    tf.AutoSize = False
    shape.Fill.Visible = False
    shape.Line.Weight = 1.0
    
    tr = tf.TextRange
    tr.Font.Name = font_name
    tr.Font.Size = font_size
    tr.Font.Bold = 1
    tr.Font.Color = 0
    tr.Text = '\n'.join('\t'.join(row) for row in grid)
    
    for p_idx in range(1, tr.Paragraphs.Count + 1):
        pf = tr.Paragraphs(p_idx).Range.ParagraphFormat
        pf.SpaceBefore = 0
        pf.SpaceAfter = 0
        pf.LineSpacingRule = 0
        pf.Alignment = 0
        pf.TabStops.ClearAll()
        for t_pt in tabs[1:]:
            pf.TabStops.Add(Position=t_pt, Alignment=0)
            
    shape.ConvertToInlineShape()
    sel.Range.InsertParagraphAfter()
    print(f'{name} -> cols={cols}, rows={num_rows}, width={box_w_pt:.1f}pt, height={box_h_pt:.1f}pt')

out_path = os.path.abspath('output/test_perfect_boxes_all.docx')
doc.SaveAs(out_path)
doc.Close()
word.Quit()
print('Finished perfectly to:', out_path)
