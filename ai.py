from gensim.models import Word2Vec

# 读取同义词语料库
sentences = []
with open('corpus.txt', 'r', encoding='utf-8') as f:
    for line in f:
        sentence = line.strip().split("\t")
        sentences.append(sentence)

# 训练Word2Vec模型
model = Word2Vec(sentences, vector_size=100, window=5, min_count=1, workers=4)

# 获取相似词
similar_words = model.wv.most_similar('星务计算机A12V电源')
print(similar_words)