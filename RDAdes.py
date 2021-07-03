import pandas as pd
import glob
from pathlib import Path
import numpy as np

# find the MarcEdit output xlsx in the downloads folder
downloads_path = str(Path.home() / "Downloads")
path1 = downloads_path
allFiles = glob.glob(path1 + "/*.xlsx")
df = pd.DataFrame()
for file_ in allFiles:
    df = pd.read_excel(file_)

# combine all the output columns for designators separating by semicolon
df['DES'] = df["100$e"].astype(str) + ';' + df["600$e"].astype(str) + ';' + df["700$e"].astype(str) + ';' + \
              df["110$e"].astype(str) + ';' + df["610$e"].astype(str) + ';' + df["710$e"].astype(str) + ';' + \
              df["111$j"].astype(str) + ';' + df["611$j"].astype(str) + ';' + df["711$j"].astype(str) + ';' + \
              df["880$e"].astype(str) + ';' + df["880$j"].astype(str)

# strip unnecessary columns
df = df[['001', '035$a', 'DES']]

# extract OCLC# from  035
df['035$a'] = df['035$a'].str.findall(r'\(OCoLC\)ocm\d{8}|\(OCoLC\)ocn\d{9}|\(OCoLC\)on\d{10}')

# remove null values - if any RDA designators end with 'nan' in the future this can be a problem
df['DES'] = df['DES'].str.replace(r'nan$|nan\;', '', regex=True)
# remove periods and commas
df['DES'] = df['DES'].str.replace(r'\.|\,', '', regex=True)

# convert DES to list
df["DES"] = df["DES"].str.split(";")

# explode multiple designators into multiple rows while retaining 001 and 035
df = df.apply(pd.Series.explode)

# drop blank rows
df['DES'] = df['DES'].replace('', np.nan)
df = df.dropna(subset=['DES'])

# remove trailing spaces
df['DES'] = df['DES'].str.rstrip()

# df for RDA relationship designators
dfRDA = pd.DataFrame(data={'TERM': ['abridger', 'abridger of', 'actor', 'actor of', 'addressee', 'addressee of',
                                    'animator', 'animator of', 'annotator', 'annotator of', 'appellant',
                                    'appellant corporate body', 'appellant corporate body of', 'appellant person',
                                    'appellant person of', 'appellee', 'appellee corporate body',
                                    'appellee corporate body of', 'appellee person', 'appellee person of',
                                    'architect', 'architect of', 'arranger of music', 'arranger of music of',
                                    'art director', 'art director of', 'artist', 'artist of', 'audio engineer',
                                    'audio engineer of', 'audio producer', 'audio producer of', 'author',
                                    'author of', 'autographer', 'autographer of', 'binder', 'binder of',
                                    'book artist', 'book artist of', 'book designer', 'book designer of',
                                    'braille embosser', 'braille embosser of', 'broadcaster', 'broadcaster of',
                                    'calligrapher', 'calligrapher of', 'cartographer', 'cartographer (expression)',
                                    'cartographer (expression) of', 'cartographer of', 'caster',
                                    'caster of', 'casting director', 'casting director of', 'censor',
                                    'censor of', 'choral conductor', 'choral conductor of',
                                    'choreographer', 'choreographer (expression)', 'choreographer (expression) of',
                                    'choreographer of', 'collection registrar', 'collection registrar of',
                                    'collector', 'collector of', 'collotyper', 'collotyper of', 'colourist',
                                    'colourist of', 'commentator', 'commentator of', 'commissioning body',
                                    'commissioning body of', 'compiler', 'compiler of', 'composer',
                                    'composer (expression)', 'composer (expression) of', 'composer of',
                                    'conductor', 'conductor of', 'consultant', 'consultant of', 'costume designer',
                                    'costume designer of', 'court governed', 'court governed of',
                                    'court reporter', 'court reporter of', 'curator', 'curator of',
                                    'current owner', 'current owner of', 'dancer', 'dancer of', 'dedicatee',
                                    'dedicatee (item)', 'dedicatee (item) of', 'dedicatee of', 'dedicator',
                                    'dedicator of', 'defendant', 'defendant corporate body',
                                    'defendant corporate body of', 'defendant person', 'defendant person of',
                                    'degree committee member', 'degree committee member of',
                                    'degree granting institution', 'degree granting institution of',
                                    'degree supervisor', 'degree supervisor of', 'depositor', 'depositor of',
                                    'designer', 'designer of', 'director', 'director of', 'director of photography',
                                    'director of photography of', 'DJ', 'DJ of', 'donor', 'donor of', 'draftsman',
                                    'draftsman of', 'dubbing director', 'dubbing director of', 'editor', 'editor of',
                                    'editor of moving image work', 'editor of moving image work of',
                                    'editorial director', 'editorial director of', 'enacting jurisdiction',
                                    'enacting jurisdiction of', 'engraver', 'engraver of', 'etcher',
                                    'etcher of', 'film director', 'film director of', 'film distributor',
                                    'film distributor of', 'film producer', 'film producer of', 'filmmaker',
                                    'filmmaker of', 'former owner', 'former owner of', 'founder of work',
                                    'founder of work of', 'honouree', 'honouree (item)', 'honouree (item) of',
                                    'honouree of', 'host', 'host institution', 'host institution of', 'host of',
                                    'illuminator', 'illuminator of', 'illustrator', 'illustrator of', 'inscriber',
                                    'inscriber of', 'instructor', 'instructor of', 'instrumental conductor',
                                    'instrumental conductor of', 'instrumentalist', 'instrumentalist of',
                                    'interviewee', 'interviewee (expression)', 'interviewee (expression) of',
                                    'interviewee of', 'interviewer', 'interviewer (expression)',
                                    'interviewer (expression) of', 'interviewer of', 'inventor', 'inventor of',
                                    'issuing body', 'issuing body of', 'judge', 'judge of',
                                    'jurisdiction governed', 'jurisdiction governed of', 'landscape architect',
                                    'landscape architect of', 'letterer', 'letterer of', 'librettist',
                                    'librettist of', 'lighting designer', 'lighting designer of',
                                    'lithographer', 'lithographer of', 'lyricist', 'lyricist of', 'make-up artist',
                                    'make-up artist of', 'medium', 'medium of', 'minute taker', 'minute taker of',
                                    'mixing engineer', 'mixing engineer of', 'moderator', 'moderator of',
                                    'music programmer', 'music programmer of', 'musical director',
                                    'musical director of', 'narrator', 'narrator of', 'on-screen participant',
                                    'on-screen participant of', 'on-screen presenter', 'on-screen presenter of',
                                    'organizer', 'organizer of', 'panelist', 'panelist of', 'papermaker',
                                    'papermaker of', 'participant in treaty', 'participant in treaty of',
                                    'performer', 'performer of', 'photographer', 'photographer (expression)',
                                    'photographer (expression) of', 'photographer of', 'plaintiff',
                                    'plaintiff corporate body', 'plaintiff corporate body of', 'plaintiff person',
                                    'plaintiff person of', 'platemaker', 'platemaker of', 'praeses',
                                    'praeses of', 'presenter', 'presenter of', 'printer', 'printer of',
                                    'printmaker', 'printmaker of', 'producer', 'producer of', 'production company',
                                    'production company of', 'production designer', 'production designer of',
                                    'programmer', 'programmer of', 'puppeteer', 'puppeteer of',
                                    'radio director', 'radio director of', 'radio producer',
                                    'radio producer of', 'rapporteur', 'rapporteur of', 'recording engineer',
                                    'recording engineer of', 'recordist', 'recordist of', 'remix artist',
                                    'remix artist of', 'researcher', 'researcher of', 'respondent', 'respondent of',
                                    'restorationist (expression)', 'restorationist (expression) of',
                                    'restorationist (item)', 'restorationist of', 'screenwriter',
                                    'screenwriter of', 'sculptor', 'sculptor of', 'seller', 'seller of', 'singer',
                                    'singer of', 'software developer', 'software developer of', 'sound designer',
                                    'sound designer of', 'speaker', 'speaker of', 'special effects provider',
                                    'special effects provider of', 'sponsoring body', 'sponsoring body of',
                                    'stage director', 'stage director of', 'storyteller', 'storyteller of',
                                    'surveyor', 'surveyor of', 'television director', 'television director of',
                                    'television producer', 'television producer of', 'transcriber',
                                    'transcriber of', 'translator', 'translator of', 'visual effects provider',
                                    'visual effects provider of', 'voice actor', 'voice actor of',
                                    'writer of added commentary', 'writer of added commentary of',
                                    'writer of added lyrics', 'writer of added lyrics of', 'writer of added text',
                                    'writer of added text of', 'writer of afterword', 'writer of afterword of',
                                    'writer of foreword', 'writer of foreword of', 'writer of introduction',
                                    'writer of introduction of', 'writer of postface',
                                    'writer of postface of', 'writer of preface', 'writer of preface of',
                                    'writer of supplementary textual content',
                                    'writer of supplementary textual content of']})

# find invalid terms
dfOut = df[~df['DES'].isin(dfRDA['TERM'])]

# output
dfOut.to_csv(path1 + "/output.csv", index=False)
