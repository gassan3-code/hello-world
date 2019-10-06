import numpy
import sys
import requests
import re
from bs4 import BeautifulSoup
import pandas as pd
from pandas import Series, DataFrame
from googletrans import Translator

class WebScrap:
    def __init__(self):
        self.keyWord = None
        self.Sent = None
        self.Shorten = None
        self.trans = None

    # 半角小文字か判定
    def islower(self,word):
        lowerReg = re.compile(r'^[a-z]*$')
        return lowerReg.match(word) is not None

    #半角英字
    def isalpha(self,word):
        alphaReg = re.compile(r'^[a-zA-Z]+$')
        return alphaReg.match(word) is not None

    # アルクで単語検索。1番最初の検索結果を返す
    def word2AlcSent(self,word):
        if self.islower(word):
            # アルクで英単語を検索する場合、頭文字が大文字である必要がある
            word = word.capitalize()
        self.keyWord = word
        base_url = "http://eow.alc.co.jp/search"
        query = {}
        query["q"] = word
        query["ref"] = "sa"
        ret = requests.get(base_url,params=query)
        text = None
        soup = BeautifulSoup(ret.content,"lxml")
        for l in soup.findAll("div",{"id":"resultsList"}):
            for m in l.findAll("div"):
                try:
                    text = m.text
                    if text is not None:
                        break;
                except:
                    pass
        self.Sent = text
        return text

    # グーグルの翻訳
    def word2GgTrancelate(self,word):
        ret = None
        if word is None:
            print("Err | Input word is invalid") 
            return ret

        translator = Translator()
        if self.isalpha(word):# en -> ja
            try: #エラー処理は暫定策、要検討
                ret = translator.translate(word, src='en', dest='ja').text
            except:
                    pass
        else: # ja -> en
            try:
                ret = translator.translate(word, src='ja', dest='en').text
            except:
                    pass
        self.trans = ret
        return ret

    #　アルクの検索結果から省略語を取得する
    def AlcSent2Shorten(self,Sent):# Shorten : 略語
        ret = None
        if Sent is None:
            print("Err | Input sentence is invalid") 
            return ret

        num = Sent.find('【略】') #指定の文字列の位置をインデックス番号で返す。見つからない場合は-1。
        if num == -1: #略語がない場合はreturn
            print("Err | there is not a shorten with ",self.keyWord) 
            return ret

        Sent = Sent[num+3:]
        numEnd = Sent.find("〕|【|◆|《")
        if num is not -1: #略語に関連しない内容を削除(正確には略語の後にくる不要部)
            ret = Sent[:numEnd]
        self.Shorten = ret
        return ret

    # エクセルファイルから入力、出力する(formatはwordListIn.xlsxに合わせる。もしくは修正して利用。)
    def mainFile(self,fin="wordListIn.xlsx",fout="wordListOut.xlsx"):
        wordList = pd.read_excel(fin)
        numMax = wordList.shape[0]
        outData = pd.DataFrame([], columns=['単語','Google翻訳','ALC省略語','ALC説明文'], index=range(numMax))

        for num in range(numMax):
            self.__init__()
            word = wordList.loc[num].values[0]
            trans = self.word2GgTrancelate(word)
            print("翻訳 : ",trans)
            sent = self.word2AlcSent(word)
            print("検索結果 : ",sent)
            shorten = self.AlcSent2Shorten(sent)
            print("略語 : ",shorten)
            addRow = [self.keyWord,self.trans,self.Shorten,self.Sent]
            outData.iloc[num,:] = addRow
        outData.to_excel(fout, sheet_name='wordList')
        print(str(fout)+"の作成が完了しました。")

    # 単語から略語を検索する。
    def mainWord(self,word):
        trans = self.word2GgTrancelate(word)
        print("翻訳 : ",trans)
        sent = self.word2AlcSent(word)
        print("検索結果 : ",sent)
        shorten = self.AlcSent2Shorten(sent)
        print("略語 : ",shorten)
        return shorten


if __name__ == "__main__":
    args = sys.argv

    # 知りたい単語名がコマンド引数がある場合は、
    # 省略語を返す。(Webで単語から省略語を検索することを想定して。必要なければ消して。)
    # pythonとWebの連携はこんな感じ？ < https://qiita.com/sandream/items/e2ecb524240d81c57ea2 >
    if 1 < len(args):
        WS = WebScrap()
        ret = WS.mainWord(args[1])
        if ret is None:
            ret = str(args[1]) + "に該当する省略語はありません。"
        print(ret)
    else:
        WS = WebScrap()
        # WS.mainWord("ボールベアリング")
        WS.mainFile() 

# 参考HP　https://oneshotlife-python.hatenablog.com/entry/2016/03/02/232705