import sys
import os
sys.path.insert(0, r'c:\Users\ACER\Desktop\RANDOM PROJECT\JSON TO DOCX CONVERTER')
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
doc = word.Documents.Add()

from renderers.exercises.boxes import BoxesRendererMixin
from renderers.common.units_and_colors import cm_to_pt
from renderers.common.typography import TypographyMixin

d = Dummy = type('Dummy', (BoxesRendererMixin, TypographyMixin), {'font_name': 'Times New Roman', 'font_size': 12.0})()
printable_width_pt = cm_to_pt(17.0)

test_boxes = [
    ('U5 Pancakes (Dialogue)', """A. Then add some yeast and a pinch of salt to the mixture. Mix with a whisk.
B. My pleasure. Enjoy your pancakes!
C. Just cook until golden then serve with some fruit or vegetables.
D. Sure. All you need is some butter, 1/2 a littre of milk, 250 grams of flour and 4 eggs.
E. Heat some butter in a frying pan and pour about 1/4 cup of the mixture into the pan at a time.
F. First, beat 4 eggs together with flour and milk.""", False),
    ('U3 Sentences (Word Bank Pipe)', "Because they like doing something useful and helping others. | Yes, it makes a better life and improves the society. | Because volunteering teaches me a lot. | Yes, I've been a volunteer teacher for Street Child Organization. | It helps you stay healthy, increases self-confidence, and makes you happy. | We can donate money or clothes via charitable organisations.", True),
    ('U2 Vocab (Word Bank Pipe)', 'active | chapped | dim | dirty | fit | healthy | tidy | tired', True),
    ('U4 Health (Word Bank Pipe)', 'allergy | exercising | playing sports | runny nose | stomachache | boating | fever | red skin | sleeping | swimming | cough | flu | relaxing | sneezing | walking | cycling | headache | running | sore throat | watching TV', True)
]

for name, raw, is_pipe in test_boxes:
    font_name = 'Times New Roman'
    font_size = 12.0
    
    if not is_pipe:
        lines = [l.strip() for l in raw.split('\n') if l.strip()]
        clean_lines = [d.strip_markup_for_measurement(l) for l in lines]
        line_widths_pt = [d.measure_text_width_pt(doc, cl, font_name, font_size, is_bold=True) for cl in clean_lines]
        max_line_w_pt = max(line_widths_pt) if line_widths_pt else 50.0
        pad_left_pt = cm_to_pt(0.20)
        pad_right_pt = cm_to_pt(0.20)
        extra_buffer_pt = cm_to_pt(0.25)
        total_pad_pt = pad_left_pt + pad_right_pt + extra_buffer_pt
        
        needed_w = max_line_w_pt + total_pad_pt + cm_to_pt(0.35)
        box_width_pt = min(printable_width_pt, max(printable_width_pt * 0.50, needed_w))
        
        shape = doc.Shapes.AddShape(5, 0, 0, box_width_pt, 50.0)
        tf = shape.TextFrame
        tf.MarginTop = 0
        tf.MarginBottom = 0
        tf.MarginLeft = pad_left_pt
        tf.MarginRight = pad_right_pt
        tf.WordWrap = -1
        tf.AutoSize = False
        shape.Fill.Visible = False
        shape.Line.Weight = 1.0
        
        tr = tf.TextRange
        tr.Font.Name = font_name
        tr.Font.Size = font_size
        tr.Font.Bold = 1
        tr.Text = '\n'.join(d.clean_box_item(l) for l in lines)
        
        for p in range(1, tr.Paragraphs.Count + 1):
            pf = tr.Paragraphs(p).Range.ParagraphFormat
            pf.SpaceBefore = 0
            pf.SpaceAfter = 0
            pf.LineSpacingRule = 0
            pf.Alignment = 0
            
        try:
            actual_lines = tr.ComputeStatistics(1)
        except Exception:
            actual_lines = len(lines)
            
        box_height_pt = (actual_lines * (font_size * 1.25)) + (font_size * 0.40) + 2.0
        shape.Height = box_height_pt
        shape.ConvertToInlineShape()
        print(f'{name} -> width={box_width_pt:.1f}pt, actual_lines={actual_lines}, height={box_height_pt:.1f}pt')
    else:
        words = [w.strip() for w in raw.split('|') if w.strip()]
        cols, grid, box_width_pt, tab_stops_pt = d.optimize_word_bank_layout(doc, words, printable_width_pt)
        pad_horiz_pt = cm_to_pt(0.20)
        
        shape = doc.Shapes.AddShape(5, 0, 0, box_width_pt, 50.0)
        tf = shape.TextFrame
        tf.MarginTop = 0
        tf.MarginBottom = 0
        tf.MarginLeft = pad_horiz_pt
        tf.MarginRight = pad_horiz_pt
        tf.WordWrap = -1 if cols == 1 else 0
        tf.AutoSize = False
        shape.Fill.Visible = False
        shape.Line.Weight = 1.0
        
        tr = tf.TextRange
        tr.Font.Name = font_name
        tr.Font.Size = font_size
        tr.Font.Bold = 1
        tr.Text = '\n'.join('\t'.join(d.clean_box_item(it) for it in chunk) for chunk in grid)
        
        for p in range(1, tr.Paragraphs.Count + 1):
            pf = tr.Paragraphs(p).Range.ParagraphFormat
            pf.SpaceBefore = 0
            pf.SpaceAfter = 0
            pf.LineSpacingRule = 0
            pf.Alignment = 0
            pf.TabStops.ClearAll()
            for t_pt in tab_stops_pt[1:]:
                pf.TabStops.Add(Position=t_pt, Alignment=0)
                
        try:
            actual_lines = tr.ComputeStatistics(1)
        except Exception:
            actual_lines = len(grid)
            
        box_height_pt = (actual_lines * (font_size * 1.25)) + (font_size * 0.40) + 2.0
        shape.Height = box_height_pt
        shape.ConvertToInlineShape()
        print(f'{name} -> width={box_width_pt:.1f}pt, actual_lines={actual_lines}, height={box_height_pt:.1f}pt')

out_path = os.path.abspath('output/test_dynamic_stats_boxes.docx')
doc.SaveAs(out_path)
doc.Close()
word.Quit()
print('SUCCESS! Saved to:', out_path)
