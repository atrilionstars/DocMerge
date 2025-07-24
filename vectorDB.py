import chromadb

chroma_client = chromadb.chromadb.Client()
collection = chroma_client.create_collection(name="test_information")
collection.add(
    ids = ["id1", "id2"],
    documents = ["Basketball is the most popular ball game in the world", "SH Company is on the verge of bankruptcy"]
)
results = collection.query(
    query_texts = ["company"],
    n_results = 1
)
print(results)