# 加载依赖的库
import matplotlib.pyplot as plt
import fenci
import os
import docx
from wordcloud import WordCloud
import jieba

# path为需要生成词云的docx文件所在的文件夹
path = '/Users/joshua/Library/Mobile Documents/com~apple~CloudDocs/学习/游记/'

# 生成所有docx文件绝对地址的列表
file_list = os.listdir(path)
for i in range(len(file_list)):
    file_list[i] = os.path.join(path,file_list[i]) 

# 读取docx中的所有文件，把它们添加到名为text的列表，同时捕捉读取异常。
# 确定只（能）读取后缀为docx的文件。有时候macbook的文档会因为保存的问题出现加了$符号的文件，也要排除掉。
text = []
n = 0
for file in file_list:
    if '$' not in file and file[-4:] == 'docx':
        try:
            doc = docx.Document(file)
            text = text+(doc.paragraphs)
        except:
            n +=1
print(n,"exception(s) occurred")

# 把text列表中所有的元素添加为data字符串中
data = ''
for para in text:
    data+=para.text

# 筛掉字符串中的标点符号
with open('/Users/joshua/Library/Mobile Documents/com~apple~CloudDocs/PYTHON/python/中文词频统计/punctuation.txt', 'r', encoding='UTF-8') as punctuationFile:
    for punctuation in punctuationFile.readlines():
        data = data.replace(punctuation[0], ' ')

# 筛掉字符串中无意义的词，比如一些连接词和副词 
with open('/Users/joshua/Library/Mobile Documents/com~apple~CloudDocs/PYTHON/python/中文词频统计/meaningless.txt', 'r', encoding='UTF-8') as meaninglessFile:
    for meaningless in meaninglessFile.readlines():
        data = data.replace(meaningless[0], '')

# 用jieba库进行分词操作
data_list = list(jieba.cut(data))

# 清除分词结果中只包含一次字的内容
for i in range(len(data_list))[::-1]:
    if len(data_list[i]) <= 1:
        data_list.pop(i)


# 生成词云图片，可以选择大小和字体
plt.subplots(figsize=(16,16))
font = r'/Library/Fonts/Microsoft/SimSun.ttf'
wordcloud = WordCloud(
                          font_path = font,
                          background_color='white',
                          width=1024,
                          height=768
                         ).generate(" ".join(data_list))
plt.imshow(wordcloud)
plt.axis('off')
plt.savefig('graph.png')
plt.show()