import os
import re
import string
import dotenv
import pandas as pd
from scipy import spatial
import ast
from openai import OpenAI
import tiktoken

dotenv.load_dotenv()
openai_api_key = os.getenv('OPENAI_API_KEY')  # Now you can access the API key
client = OpenAI(api_key=openai_api_key)
token_budget = 2000
embedding_model = "text-embedding-3-large"
gpt4 = "gpt-4-0125-preview"
gpt3 = "gpt-3.5-turbo-0125"
GPT_MODEL = gpt3
# Set display options to show all columns and rows
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

filepath = 'text/NewOrleansCodes.xlsx'
embedding_path = 'downloads/cityembeddings.csv'

# create_embedding_df(filepath, embedding_path)
print('loading embeddings...')
df = pd.read_csv(embedding_path)

# convert embeddings from CSV str type back to list type
df['embedding'] = df['embedding'].apply(ast.literal_eval)
print('embeddings loaded')
print(df.head())


# search function
def strings_ranked_by_relatedness(
    query: str,
    df: pd.DataFrame,
    relatedness_fn=lambda x, y: 1 - spatial.distance.cosine(x, y),
    top_n: int = 100
) -> tuple[list[str], list[float]]:
    """Returns a list of strings and relatednesses, sorted from most related to least."""
    query_embedding_response = client.embeddings.create(
        model=embedding_model,
        input=query,
    )
    query_embedding = query_embedding_response.data[0].embedding
    strings_and_relatednesses = [
        (row["text"], relatedness_fn(query_embedding, row["embedding"]))
        for i, row in df.iterrows()
    ]
    strings_and_relatednesses.sort(key=lambda x: x[1], reverse=True)
    strings, relatednesses = zip(*strings_and_relatednesses)
    return strings[:top_n], relatednesses[:top_n]


def relatedness_score(text, _df):
    # examples
    strings, relatednesses = strings_ranked_by_relatedness(text, _df, top_n=3)
    for string, relatedness in zip(strings, relatednesses):
        print(f"{relatedness=:.3f}")
        print(string)


def remove_stuff(text: str) -> str:
    """Remove punctuation (except in URLs), newline, tab characters, and large spaces."""
    # Pattern to identify URLs
    url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    # Find all URLs using the pattern
    urls = re.findall(url_pattern, text)
    # Replace URLs with a placeholder to avoid altering them
    placeholder = "URL_PLACEHOLDER"
    for url in urls:
        text = text.replace(url, placeholder)

    # Remove punctuation
    text = re.sub(r'[^\w\s]', '', text)
    # Remove newline and tab characters
    text = re.sub(r'[\n\t]+', '', text)
    # Remove large spaces (5 or more spaces)
    text = re.sub(r' {5,}', ' ', text)

    # Restore URLs from placeholders
    for url in urls:
        text = text.replace(placeholder, url, 1)

    return text

"""def remove_stuff(text: str) -> str:
    # Create a translation table that maps each punctuation character to None
    translator = str.maketrans('', '', string.punctuation)
    # Use the translate function to remove punctuation
    text = text.translate(translator)
    # Remove newline and tab characters
    text = re.sub(r'[\n\t]+', '', text)
    return text"""


def get_embedding(text_to_embed):
    text_to_embed = remove_stuff(text_to_embed)
    # Embed a line of text
    response = client.embeddings.create(
        model=embedding_model,
        input=[text_to_embed]
    )
    # Extract the AI output embedding as a list of floats
    embedding = response.data[0].embedding
    print(f"---\nEmbedding: {embedding} \nText: {text_to_embed}")

    return embedding


def create_embedding_df(excel_path, embedding_path):
    # Create a new DataFrame to store the text and its embedding
    _df = pd.DataFrame()

    # Read the Excel file and set row 0 as the header
    review_df = pd.read_excel(excel_path, header=None)
    # First, let's set the correct headers if they are present in the second row (row index 1)
    review_df.columns = review_df.iloc[1]  # This sets the second row as the header
    review_df = review_df.drop([0, 1])  # This drops the first two rows which are now redundant

    # Now concatenate all the text in each row into a single cell in a new column called 'text'
    _df['text'] = review_df.apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
    _df.reset_index(drop=True, inplace=True)
    # Display the DataFrame with the concatenated column
    print(_df.head())
    # _df = _df.sample(10)
    _df["embedding"] = _df["text"].astype(str).apply(get_embedding)
    print(_df.head())

    _df.to_csv(embedding_path, index=False)






def num_tokens(text: str, model: str = GPT_MODEL) -> int:
    """Return the number of tokens in a string."""
    encoding = tiktoken.encoding_for_model(model)
    return len(encoding.encode(text))


def query_message(
    query: str,
    df: pd.DataFrame,
    model: str,
    token_budget: int
) -> str:
    """Return a message for GPT, with relevant source texts pulled from a dataframe."""
    strings, relatednesses = strings_ranked_by_relatedness(query, df)
    introduction = 'Use the New Orleans Code of Ordinaces provided below to answer the subsequent question. ' \
                   'If the answer cannot be found in the articles, say so, ' \
                   'then try to answer the question as best you can anyways'
    question = f"\n\nQuestion: {query}"
    message = introduction
    for string in strings:
        next_article = f'\n\nNew Orleans Ordinance Codes:\n"""\n{string}\n"""'
        if (
            num_tokens(message + next_article + question, model=model)
            > token_budget
        ):
            break
        else:
            message += next_article
    return message + question


def ask(
    query: str,
    df: pd.DataFrame = df,
    model: str = GPT_MODEL,
    token_budget: int = token_budget,
    print_message: bool = False,
) -> str:
    """Answers a query using GPT and a dataframe of relevant texts and embeddings."""
    message = query_message(query, df, model=model, token_budget=token_budget)
    if print_message:
        print(message)
    messages = [
        {"role": "system", "content": "You answer questions about the New Orleans City Ordinances"
                                      "Respond in a casual and easily understandable manner."
                                      "If applicable reference specific articles, and provide the links. ."},
        {"role": "user", "content": message},
    ]
    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0
    )
    response_message = response.choices[0].message.content
    if print_message is True:
        print(response_message)
        completion_tokens = response.usage.completion_tokens
        completion_price = completion_tokens * (0.0015 / 1000)
        print(f"Number of completion tokens: {completion_tokens} ${completion_price}\n")

        prompt_tokens = response.usage.prompt_tokens
        prompt_price = prompt_tokens * (0.0005 / 1000)
        print(f"Number of prompt tokens: {prompt_tokens} ${prompt_price}\n")

        tokens_used = response.usage.total_tokens
        total_price = completion_price + prompt_price
        print(f"\nNumber of tokens used: {tokens_used}\nTotal Price: ${total_price}\n")
    return response_message


#relatedness_score("what happens if my license is invalid", df)

#answer = ask("What ordinances reference religion", print_message=True)
#print(answer)
