import soundcard as sc
import soundfile as sf
import threading
import speech_recognition as sr
import datetime as dt
import time
import pandas as pd

endFlag = False #記録終了フラグ
lock = threading.RLock() #ロックオブジェクトの作成
r = sr.Recognizer() #言葉を認識するオブジェクト
result_list = [] #エクセルに出力するデータを格納するリスト
output_file_name = "out.wav"    # 出力するファイル名
samplerate = 48000              # サンプリング周波数 [Hz]
record_sec = 20                 # 録音する時間 [秒] (1時間しか対応していない)

def record():
  print("記録を開始しました。終了するには「Ctrl + C」を押してください")
  with sc.get_microphone(id=str(sc.default_speaker().name), include_loopback=True).recorder(samplerate=samplerate) as mic:
      data = mic.record(numframes=samplerate*record_sec)

      # マルチチャンネルで保存したい場合は、"data=data[:, 0]"を"data=data"に変更
      sf.write(file=output_file_name, data=data[:, 0], samplerate=samplerate)

  transcription()

def transcription():
  with sr.AudioFile("out.wav") as source:
    while True:
        audio_data = r.listen(source)
        flac_data = audio_data.get_flac_data(
            convert_rate=None if audio_data.sample_rate >= 8000 else 8000,
            convert_width=2
        )
        print("len: {} ← 10MB に近づくとエラーになりやすいので、サイズを見てパメータをいじる".format(len(flac_data)))
        try:
            now = dt.datetime.today()
            result_list.append(now, r.recognize_google(audio_data, language='ja-JP'))
        except sr.UnknownValueError:
            print("Oops! Didn't catch that") # セグメントによっては音声認識できない、無視
        if len(flac_data) == 0: # ファイルの終わり
          fileName = str(now)+"_会議メモ.xlsx" #保存するファイル名
          df = pd.DataFrame(result_list,columns=['時刻', '内容'])#列名
          with pd.ExcelWriter('test.xlsx') as writer:
            df.to_excel(writer,index=False)#エクセルファイルに書き出し
          print(fileName+"という名前で保存しました。")
          break
        time.sleep(1) # ちょっと投げすぎを気にしている

record()