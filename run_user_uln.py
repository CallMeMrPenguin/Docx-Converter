import os
from uln_compiler import ULNCompiler

uln_input = """[H1] TEST FOR UNIT 1

[P0] I. Choose the word whose underlined part is pronounced differently from the others.
[P1] 1. A. [e]{u}xciting  B. [e]{u}xcellent  C. [e]{u}xperience  D. [e]{u}xpensive
[P1] 2. A. hobb[y]{u}  B. c[y]{u}cle  C. b[y]{u}e  D. wh[y]{u}
[P1] 3. A. p[o]{u}ttery  B. c[o]{u}llect  C. mel[o]{u}dy  D. m[o]{u}nopoly
[P1] 4. A. h[ea]{u}rd  B. b[i]{u}rd  C. w[o]{u}rld  D. pict[u]{u}re
[P1] 5. A. bird-wat[ch]{u}ing  B. [ch]{u}ildren  C. s[ch]{u}ool  D. [ch]{u}allenge

[P0] III. Choose the best answer a, b, c, or d to complete the sentence.
[P0] 1. My father can make beautiful pieces of art ______ empty eggshells.
[P1] A. of  B. from  C. in  D. into
[P0] 2. Why don't you take ______ a new hobby?
[P1] A. up  B. in  C. over  D. after
[P0] 3. Collecting cars is a(n) ______ hobby. It costs a lot of money.
[P1] A. interesting  B. cheap  C. expensive  D. unusual
[P0] 4. More people are ______ birds today than ever before.
[P1] A. seeing  B. looking  C. hearing  D. watching

[P0] IV. Complete the sentences with the correct form or tense of the verb play, go, do or collect.
[P0] 1. ______ Alex ______ computer games at the weekend?
[P0] 2. We ______ camping in Dam Sen Park yesterday.
[P0] 3. ______ you ______ coins some day in the future?
[P0] 4. Do you want ______ mountain biking with me this weekend?

[P0] VI. There is one mistake in each sentence. Underline and correct the mistake.
[TAB2] [P0] 1. Nam is my classmates. He watches TV every night.  ______
[TAB2] [P0] 2. I think collecting stamps are interesting.  ______
[TAB2] [P0] 3. My dad cooks very good. He loves preparing meals for our family.  ______
[TAB2] [P0] 4. I enjoy to ride my bike to school.  ______

[P0] VIII. Choose the word which best fits each gap.
[QUOTE]
Many people (1)______ crafting with paper. The materials are readily available and don't cost much; and no super special talents are needed. Anyone (2)______ be a paper crafter.

There are many different paper craft techniques. Origami is one of ancient techniques developed in Japan where squares of paper are (3)______ and formed into various objects such as flowers, animals, and boxes. Card (4)______ is also a favourite paper craft technique. Birthday cards are the most popular greeting cards, followed by Christmas cards. Receiving a (5)______ card is a special gift, because of the time and effort someone spent making it. It lets the recipient know just how much you care (6)______ them.
[/QUOTE]
[P1] 1. A. enjoy  B. decide  C. want  D. learn
[P1] 2. A. must  B. should  C. can  D. will
[P1] 3. A. wrapped  B. folded  C. torn  D. taken

[P0] X. Write sentences, using the cues given.
[P0] 1. I/ enjoy/ play/ sports/ because/ it/ good/ health
[P1] __________________________________________________________________________________
[P0] 2. your children/ go/ camp/ every summer holiday?
[P1] __________________________________________________________________________________
[P0] 3. I/ think/ photography/ can/ expensive hobby
[P1] __________________________________________________________________________________

[H1] A. PHONETICS

[P0] I. Underline the sound /f/ and circle the sound /v/.
[P1] fun    fine    coffee    over    graph
[P1] phone    brave    verb    stuff    clever
[P1] enough    laughing    leaf    leave    vitamin
[P1] view    few    valley    save    valley

[P0] II. Say the sentences out loud. Then write the words with the sound /f/ and /v/ in the table.
[P0] 1. I feel so bad. Maybe I should take a rest for some minutes.
[P0] 2. His wife is laughing at the picture of the knight on the floor.
[P0] 3. Live our life and hold our fate.
[P0] 4. Which is the best movie in Fast and Furious series?
[P0] 5. The invitation cards are beautiful and creative.
[P0] 6. What animals have the rough skin? - Elephants, frogs, etc.
[TABLE]
[TH] /f/ | /v/
[TR] ______ | ______
[TR] ______ | ______
[TR] ______ | ______
[TR] ______ | ______
[TR] ______ | ______
[/TABLE]

[P0] I. Look at the pictures and name the activities.
[BOX] cooking | horse riding | collecting coins | jogging | making models | building dollhouses | gardening | doing yoga [/BOX]
[PIC_GRID]
1. [PIC: "A child making models" | size:small] | 2. [PIC: "A child with a dollhouse" | size:small] | 3. [PIC: "Hands gardening a plant" | size:small] | 4. [PIC: "A person cooking" | size:small]
5. [PIC: "A person doing yoga outdoors" | size:small] | 6. [PIC: "Two people jogging" | size:small] | 7. [PIC: "A close-up of collecting coins" | size:small] | 8. [PIC: "A person horse riding" | size:small]
[/PIC_GRID]

[P0] A. Match each word with its meaning.
[TAB2] [P1] 1. rest  A. an activity that you enjoy doing
[TAB2] [P1] 2. recreation  B. a strong feeling of excitement and interest in something
[TAB2] [P1] 3. interest  C. a feeling of happiness, enjoyment, or satisfaction
[TAB2] [P1] 4. enthusiasm  D. a period of relaxing, sleeping or doing nothing
[TAB2] [P1] 5. relaxation  E. the fact of people doing things for enjoyment
[TAB2] [P1] 6. pleasure  F. a way of resting and enjoying yourself
"""

compiler = ULNCompiler({
    "font_name": "Times New Roman",
    "font_size": 12.0,
    "margin_top": 2.0,
    "margin_bottom": 2.0,
    "margin_left": 3.0,
    "margin_right": 1.5,
    "enable_page_numbers": True
})

output_file = os.path.join(compiler.output_dir, "unit_1_test_formatted.docx")
print("Compiling ULN input to Word document with real-time visual editing...")
compiled_path = compiler.compile(uln_input, output_file, keep_open=True)
print(f"Compilation finished! Word is open and interactive on screen at: {compiled_path}")
