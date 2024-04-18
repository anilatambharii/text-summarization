import nltk
nltk.download('punkt')

import docx2txt
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer

# Load the Word document
text = docx2txt.process("Text.docx")

# Initialize a parser with the text
parser = PlaintextParser.from_string(text, Tokenizer("english"))

# Initialize an LSA (Latent Semantic Analysis) summarizer
summarizer = LsaSummarizer()

# Generate the summary with 100 words
summary = summarizer(parser.document, 5)

# Print the summary
for sentence in summary:
    print(sentence)
