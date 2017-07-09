using CommandLine;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.Speech.V1;
using Grpc.Auth;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
namespace VCC2
{
    class Program
    {
        static string program = "";
        static string command = "";
        static string voice = "";
        //static System.Diagnostics.Process process = new System.Diagnostics.Process();
        //static System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
        [DllImport("user32.dll")]
        public static extern void keybd_event(uint vk, uint scan, uint flags, uint extraInfo);

        static string MakeQuery(string searchword)
        {
            string query = null;
            string[] list_searchWord = searchword.Split(' ');
            foreach (var word in list_searchWord)
            {
                query += word + "+";
            }
            return query = query.Substring(0, query.Length - 1);

        }
        static void SearchWeb(string searchword,string web = "")
        {
            string query = MakeQuery(searchword);
            string url = null;
            if (web.Equals("네이버"))
            {
                url= "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=" + query;
            }else
            {
                url = "https://www.google.co.kr/search?q=" + query;
            }
            url = "\"" + url + "\"";
            System.Diagnostics.Process.Start("explorer", url);
        }

        static void InitCommand()
        {
            //startInfo.FileName = "CMD.exe";
            //startInfo.WorkingDirectory = @"D:\";

            //// CMD 창 띄우기 -- true(x), false(띄우기)
            //startInfo.CreateNoWindow = false;
            //startInfo.UseShellExecute = false;
            //// CMD 데이터 받기
            //startInfo.RedirectStandardOutput = true;
            //// CMD 데이터 보내기
            //startInfo.RedirectStandardInput = true;
            //// CMD 오류내용 받기
            //startInfo.RedirectStandardError = true;

            //process.StartInfo = startInfo;
            
        }
        static void ExecCommand()
        {
           
            if (command.Equals("open")&&!program.Equals(""))
            {
                //process.Start("control");
                //// CMD에 보낼 명령어를 입력.
                //process.StandardInput.Write();
                //process.StandardInput.Close();
                /* 탐색기 기능 인수 구분 구현 예정*/
                if(program.Equals("internet"))    
                    System.Diagnostics.Process.Start("explorer","http://google.com");
                else
                    System.Diagnostics.Process.Start(program);
            }

            // 결과 값을 리턴 받기
            //string resultValue = process.StandardOutput.ReadToEnd();
            //process.WaitForExit();
            //process.Close();

           
        }
        static void CheckCommand(string voice)
        {
            if (voice.Contains("인터넷"))
                program = "internet";
            else if (voice.Contains("메모장"))
                program = "notepad";
            else if (voice.Contains("탐색기"))
                program = "explorer";
            else if (voice.Contains("제어판"))
                program = "control";
            else if (voice.Contains("캡쳐도구"))
                program = "snippingtool";
            else if (voice.Contains("스크린샷"))
            {
                /* VK_LWIN = 0x5B , VK_SNAPSHOT = 0x2C
                 * WIN + PRINT_SCREEN = 전체 캡쳐 후 사진 - 스크린샷 폴더에 저장
                 * flags : 0x00 (DOWN) 0x02 (UP) */
                keybd_event(0x5B, 0, 0x00, 0);
                keybd_event(0x2C, 0, 0x00, 0);
                keybd_event(0x5B, 0, 0x02, 0);
                keybd_event(0x2C, 0, 0x02, 0);
            }else if (voice.Contains("검색"))
            {
                command = "search";
                // 검색어 추출
                string searchword = voice.Substring(0,voice.IndexOf("검색"));
                if (searchword == null)
                {
                    Console.WriteLine("검색할 검색어가 없습니다.");
                    return;
                }
                // 공백 제거
                searchword = searchword.Trim();
                SearchWeb(searchword);
            }
            if (voice.Contains("실행")|| voice.Contains("켜"))
                command = "open";
        }
        // [START speech_streaming_mic_recognize]
        static async Task<object> StreamingMicRecognizeAsync(int seconds)
        {
            if (NAudio.Wave.WaveIn.DeviceCount < 1)
            {
                Console.WriteLine("No microphone!");
                return -1;
            }
            var speech = SpeechClient.Create();
            var streamingCall = speech.StreamingRecognize();
            // Write the initial request with the config.
            await streamingCall.WriteAsync(
                new StreamingRecognizeRequest()
                {
                    StreamingConfig = new StreamingRecognitionConfig()
                    {
                        Config = new RecognitionConfig()
                        {
                            Encoding =
                            RecognitionConfig.Types.AudioEncoding.Linear16,
                            SampleRateHertz = 16000,
                            LanguageCode = /*LanguageCodes.Korean.SouthKorea,*/"en",
                        },
                        InterimResults = true,
                    }
                });
            // Print responses as they arrive.
            Task printResponses = Task.Run(async () => 
            {
                while (await streamingCall.ResponseStream.MoveNext(
                    default(CancellationToken)))
                {
                    foreach (var result in streamingCall.ResponseStream
                        .Current.Results)
                    {
                        foreach (var alternative in result.Alternatives)
                        {
                            Console.WriteLine(alternative.Transcript);
                            voice = alternative.Transcript;
                        }
                    }
                }
            });
            // Read from the microphone and stream to API.
            object writeLock = new object();
            bool writeMore = true;
            var waveIn = new NAudio.Wave.WaveInEvent();
            waveIn.DeviceNumber = 0;
            waveIn.WaveFormat = new NAudio.Wave.WaveFormat(16000, 1);
            waveIn.DataAvailable +=
                (object sender, NAudio.Wave.WaveInEventArgs args) =>
                {
                    lock (writeLock)
                    {
                        if (!writeMore) return;
                        streamingCall.WriteAsync(
                            new StreamingRecognizeRequest()
                            {
                                AudioContent = Google.Protobuf.ByteString
                                    .CopyFrom(args.Buffer, 0, args.BytesRecorded)
                            }).Wait();
                    }
                };
            waveIn.StartRecording();
            Console.WriteLine("Speak now.");
            await Task.Delay(TimeSpan.FromSeconds(seconds));
            // Stop recording and shut down.
            waveIn.StopRecording();
            lock (writeLock) writeMore = false;
            await streamingCall.WriteCompleteAsync();
            await printResponses;
            CheckCommand(voice);
            ExecCommand();
            return 0;
        }
        // [END speech_streaming_mic_recognize]

        public static int Main(string[] args)
        {
            //InitCommand();
            return (int)StreamingMicRecognizeAsync(5).Result; 
        }
    }
}
