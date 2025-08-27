import math
import pyaudio
import numpy as np
import time
import threading 
# sample_rate = 5000
# freq = 650.0
# duration = 2
# volume = 0.5



class AlarmPlayer:
    def __init__(self, freq = 1350.0, duration=1.5, sample_rate=5000, volume=0.5, ):
        self.freq = freq
        self.duration = duration
        self.sample_rate=sample_rate
        self.volume=  volume
        self._stop_event = threading.Event()
        self.thread=None
        self.pa = pyaudio.PyAudio()

    def alarm(self):
        def run():
            t = np.linspace(0, self.duration, int(self.sample_rate*self.duration), endpoint=False)
            samples = self.volume * np.sin(2*np.pi*self.freq*t)

            audio_data = samples.astype(np.float32)

            stream = self.pa.open(format=pyaudio.paFloat32,
                            channels=1,
                            rate = self.sample_rate, 
                            output=True)
            chunk_size = 500 # samples per chunk
            for i in range(0, len(audio_data), chunk_size):
                if self._stop_event.is_set():
                    break
                chunk = audio_data[i:i+chunk_size].tobytes()
                stream.write(chunk)
            stream.close()
        self.thread = threading.Thread(target=run)
        self.thread.start()

    def cancel(self):
        self._stop_event.set()
        if self.thread:
            self.thread.join()

    def chirp(self, freq=650.0, chirp_dur=0.2, reps=2):
        """ does a nice little chirp """
        wait_time = 0.1
        volume=0.5

        t = np.linspace(0, chirp_dur, int(self.sample_rate*chirp_dur), endpoint=False)
        samples = volume * np.sin(2*np.pi*freq*t)

        audio_data = samples.astype(np.float32).tobytes()

        stream = self.pa.open(format=pyaudio.paFloat32,
                        channels=1,
                        rate = self.sample_rate, 
                        output=True)

        for _ in range(reps):
            stream.write(audio_data)
            # stream.stop_stream()
            time.sleep(wait_time)
        stream.close()

    def __del__(self):
        """ just terminates the session """
        self.pa.terminate()


# # calling alarm should capture its instance 

# alarm_player = AlarmPlayer()
# # alarm_player.alarm()
# # while True:
# #     print('hah')
# alarm_player.chirp()

# # time.sleep(.4)
# # alarm_player.cancel()
