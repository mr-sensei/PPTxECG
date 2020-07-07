#####Some constants
#
# Word length Comparisons
# Compares slide word counts to other types of text.
# Ideally slides are to be presented - they should not be wordier than comparable informative texts. 
WLC = [
    (200000, "Tolstoy bows to you. This has more words than War and Peace."),
    (150000, "This has more text than most novels this century."),
    (88942, "Longer than Orwell's 1984."),
    (60000, "Exceeds many PHD theses."),
    (50000, "Huge. This is longer than many novels."),
    (40000, "Huge. This exceeds a Masters thesis."),
    (17500, "Very long. This is longer than many novellas."),
    (7500, "Long. This qualifies as a novelette."),
    (5000, "Long. This has over 500 words.")
]

# This looks at words per minute. The developer has seen slides with higher word counts than many delegates
# could read in the time alloted. Given that the presenter is also adding information, this essentially makes
# the learning inaccessible to some - in extreme cases this could fall foul of equality legislation. In most
# cases, it makes the learer experience dull and overwhelming.
WPM = [
    (600, "Impossible. Only rappers need apply."),
    (220, "Unteachable. Most can read around this speed. Spoken speech is much slower."),
    (150, "Unteachable. Fast for a script, impossible for slides. Reseach shows faster than this is uncomfortable."),
    (100, "Unteachable. Okay for a script. Training videos (e.g. CBT) speak around this rate."),
    (30, "Barely teachable. These are busy slides."),
    # Anything above this is a script/book - many trainers will panic and start reading out slides.
    # Better trainers will cut content. In which case, why was it even there?
    (15, "Fast. Delegate may be trying to read text over trainer's shoulder."),
    (8, "Okay. Delegates can take in info in a few seconds, then focus on trainer. Good.")
]

#Slide word counts
MANUAL = 200            # Booklike: a dense page of text.
BOOKLIKE = 120          # Booklike: more suited to a book.
INSANE = 90             # Death by Powerpoint. Have an ECG on hand.
EXTREME = 70            # Is that snoring?
HIGH = 50               # Getting high. Some slides here may be okay.
RULE_OF_33 = 33         # Can be okay if slides are well designed
FIVE_BY_FIVE = 25       # Absolute upper limit of good.
FOUR_BY_FOUR = 16       # Okay. Some guidance recommends.
STEVE_JOBS = 8          # Great! 190 point text! Slides support you.

# Excluding outlying 'speed readers,' adult reading speeds tend to vary from around 100~600wpm, with an average around 200wpm.
# However, this also tends to depend on the entropy and familiarity of the text - e.g. clearly written, redundant prose, such
# as magazines or self-help books tend to be much faster (i.e. 3-10x) to read than technical, jargon-filled or unfamiliar writing.
# Given that this is for use in education, we'll go low - assuming some slower readers and tough content.

SLOWEST_READER_WPM = 100

# Comments for slide output.
Word_Counts_Comments = [
    (0, "No text."),
    (STEVE_JOBS,"Excellent. Minimal word count means students can focus on trainer input."),
    (FOUR_BY_FOUR,"Good. Reasonable amount of content."),
    (FIVE_BY_FIVE,"Okay. The amount of content may distract student attention."),
    (RULE_OF_33,"Okay. A lot of info here. Okay for a few info-heavy/technical slides, but you don't want all slides like this."),
    (HIGH,"Caution: getting quite text heavy. Perhaps not yet 'Death by PowerPoint' territory, but you can see it from here. "), 
    (EXTREME,"Caution: getting very text heavy. Dipping well into 'Death by PowerPoint territory. "), 
    (INSANE,"Warning: this slide is very text heavy and probably not 'presentatable.'"),
    (BOOKLIKE,"Warning: this is more like a page from a book than a presentation."),
    (MANUAL,"Extreme. This is essentially a page of dense text from a reference book."),
]