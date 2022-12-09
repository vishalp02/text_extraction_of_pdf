import spacy
import pandas as pd

extraction_df = pd.read_excel('Extracted_sentences.xlsx')

# python -m spacy download en_core_web_sm
nlp = spacy.load('en_core_web_sm')


verbs = []

for i,row in extraction_df.iterrows():
    doc = nlp(row['Sentences'])

    verbs.append([token.lemma_ for token in doc if token.pos_ == 'VERB'])

verbs_df = pd.DataFrame()

verb_data=[]

for j in range(len(verbs)):
    verb_data.append(', '.join(verbs[j]))
    
verbs_df['Verb-entities'] = verb_data

re = pd.concat([extraction_df, verbs_df], axis=1, join='inner')

re.to_excel('Final_Result.xlsx', index=False)