import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import openai
import io

# Set your OpenAI key here
openai.api_key = "sk-abcdef1234567890abcdef1234567890abcdef12"  # Replace this with your key

st.set_page_config(page_title="Excel Chat Assistant", layout="centered")
st.title("🤖 Excel Chat Assistant")
st.write("Upload your Excel file and ask questions in plain English!")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

def normalize_column_names(df):
    # Convert column names to strings before applying strip() and replace()
    df.columns = [str(col).strip().lower().replace(" ", "").replace("/", "") for col in df.columns]
    return df

def show_chart(df, question):
    fig, ax = plt.subplots()
    # Simple demo logic: look for specific keywords
    if "by" in question:
        parts = question.split("by")
        target = parts[0].strip().split()[-1]
        group_by = parts[1].strip().split()[0]
        if target in df.columns and group_by in df.columns:
            sns.barplot(x=group_by, y=target, data=df, ax=ax)
            plt.xticks(rotation=45)
            st.pyplot(fig)
        else:
            st.write("Cannot find matching columns for chart.")
    else:
        st.write("Please include 'by' in your question for charting.")

def ask_openai(df, question):
    sample_data = df.head(10).to_csv(index=False)
    prompt = f"""
You are a helpful data assistant. You will be given a sample of a dataset and a question. Answer the question accurately.

DATA:
{sample_data}

QUESTION:
{question}

ANSWER:
"""
    # New API call for OpenAI (v1.0)
    response = openai.completions.create(
        model="gpt-3.5-turbo",  # Or use gpt-3.5-turbo if needed
        prompt=prompt,
        temperature=0.7,
        max_tokens=150
    )
    return response['choices'][0]['text'].strip()

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df = normalize_column_names(df)
    st.subheader("Data Preview")
    st.dataframe(df.head())

    user_query = st.text_input("Ask a question about your data")

    if user_query:
        if any(word in user_query.lower() for word in ["chart", "plot", "graph"]):
            show_chart(df, user_query.lower())
        else:
            with st.spinner("Thinking..."):
                answer = ask_openai(df, user_query)
                st.write("Answer:")
                st.write(answer)
else:
    st.info("Please upload an Excel (.xlsx) file to get started.")