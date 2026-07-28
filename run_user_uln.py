import os
from uln_compiler import ULNCompiler

uln_input = """[H1] TEST FOR UNIT 1

[P0] I. Choose the word whose underlined part is pronounced differently from the others.
[P1] 1. a. [e]{u}xciting \t b. [e]{u}xcellent \t c. [e]{u}xperience \t d. [e]{u}xpensive
[P1] 2. a. hobb[y]{u} \t b. c[y]{u}cle \t c. b[y]{u}e \t d. wh[y]{u}
[P1] 3. a. p[o]{u}ttery \t b. c[o]{u}llect \t c. mel[o]{u}dy \t d. m[o]{u}nopoly
[P1] 4. a. h[ea]{u}rd \t b. b[i]{u}rd \t c. w[o]{u}rld \t d. pict[u]{u}re
[P1] 5. a. bird-wat[ch]{u}ing \t b. [ch]{u}ildren \t c. s[ch]{u}ool \t d. [ch]{u}allenge

[P0] III. Choose the best answer a, b, c, or d to complete the sentence.
[P0] 1. My father can make beautiful pieces of art _____ empty eggshells.
[P1] a. of \t b. from \t c. in \t d. into
[P0] 2. Why don't you take _____ a new hobby?
[P1] a. up \t b. in \t c. over \t d. after
[P0] 3. Collecting cars is a(n) _____ hobby. It costs a lot of money.
[P1] a. interesting \t b. cheap \t c. expensive \t d. unusual
[P0] 4. More people are _____ birds today than ever before.
[P1] a. seeing \t b. looking \t c. hearing \t d. watching

[P0] IV. Complete the sentences with the correct form or tense of the verb play, go, do or collect.
[P0] 1. _____ Alex _____ computer games at the weekend?
[P0] 2. We _____ camping in Dam Sen Park yesterday.
[P0] 3. _____ you _____ coins some day in the future?
[P0] 4. Do you want _____ mountain biking with me this weekend?

[P0] VI. There is one mistake in each sentence. Underline and correct the mistake.
[TAB2] [P0] 1. Nam is my classmates. He watches TV every night. \t _____
[TAB2] [P0] 2. I think collecting stamps are interesting. \t _____
[TAB2] [P0] 3. My dad cooks very good. He loves preparing meals for our family. \t _____
[TAB2] [P0] 4. I enjoy to ride my bike to school. \t _____

[P0] VIII. Choose the word which best fits each gap.
[QUOTE]
Many people (1)_____ crafting with paper. The materials are readily available and don't cost much; and no super special talents are needed. Anyone (2)_____ be a paper crafter.
There are many different paper craft techniques. Origami is one of ancient techniques developed in Japan where squares of paper are (3)_____ and formed into various objects such as flowers, animals, and boxes. Card (4)_____ is also a favourite paper craft technique. Birthday cards are the most popular greeting cards, followed by Christmas cards. Receiving a (5)_____ card is a special gift, because of the time and effort someone spent making it. It lets the recipient know just how much you care (6)_____ them.
[/QUOTE]
[P1] 1. a. enjoy \t b. decide \t c. want \t d. learn
[P1] 2. a. must \t b. should \t c. can \t d. will
[P1] 3. a. wrapped \t b. folded \t c. torn \t d. taken

[P0] X. Write sentences, using the cues given.
[P0] 1. I/ enjoy/ play/ sports/ because/ it/ good/ health
[P1] _____
[P0] 2. your children/ go/ camp/ every summer holiday?
[P1] _____
[P0] 3. I/ think/ photography/ can/ expensive hobby
[P1] _____

[H1] A. PHONETICS

[P0] I. Underline the sound /f/ and circle the sound /v/.
[P0] fun \t fine \t coffee \t over \t graph
[P0] phone \t brave \t verb \t stuff \t clever
[P0] enough \t laughing \t leaf \t leave \t vitamin
[P0] view \t few \t valley \t save \t valve

[P0] II. Say the sentences out loud. Then write the words with the sound /f/ and /v/ in the table.
[P0] 1. I feel so bad. Maybe I should take a rest for some minutes.
[P0] 2. His wife is laughing at the picture of the knight on the floor.
[P0] 3. Live our life and hold our fate.
[P0] 4. Which is the best movie in Fast and Furious series?
[P0] 5. The invitation cards are beautiful and creative.
[P0] 6. What animals have the rough skin? - Elephants, frogs, etc.
[TAB2] /f/ \t /v/
[TAB2] _____ \t _____
[TAB2] _____ \t _____
[TAB2] _____ \t _____

[P0] I. Look at the pictures and name the activities.
[BOX] cooking \t horse riding \t collecting coins \t jogging \t making models \t building dollhouses \t gardening \t doing yoga [/BOX]
[TAB2] [PIC: "Boy assembling a model kit" | pos:center | size:small] \t [PIC: "Girl standing next to a dollhouse" | pos:center | size:small] \t [PIC: "Hands planting a small plant in soil" | pos:center | size:small] \t [PIC: "Woman cooking food in kitchen" | pos:center | size:small]
[TAB2] [P1] 1. _____ \t [P1] 2. _____ \t [P1] 3. _____ \t [P1] 4. _____
[TAB2] [PIC: "Woman doing yoga pose outdoors" | pos:center | size:small] \t [PIC: "Two people jogging on path" | pos:center | size:small] \t [PIC: "Coins and small objects collected on table" | pos:center | size:small] \t [PIC: "Woman riding a white horse in field" | pos:center | size:small]
[TAB2] [P1] 5. _____ \t [P1] 6. _____ \t [P1] 7. _____ \t [P1] 8. _____

[P0] A. Match each word with its meaning.
[TAB2] [P1] 1. rest \t a. an activity that you enjoy doing
[TAB2] [P1] 2. recreation \t b. a strong feeling of excitement and interest in something
[TAB2] [P1] 3. interest \t c. a feeling of happiness, enjoyment, or satisfaction
[TAB2] [P1] 4. enthusiasm \t d. a period of relaxing, sleeping or doing nothing
[TAB2] [P1] 5. relaxation \t e. the fact of people doing things for enjoyment
[TAB2] [P1] 6. pleasure \t f. a way of resting and enjoying yourself
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
