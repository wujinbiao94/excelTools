import jieba

def wordFrequencyCount(message) :
    words = jieba.cut(message)
    return words
