using CommandLine;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.Speech.V1;
using Grpc.Auth;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using System.Text;
using System.Diagnostics;

namespace VCC2
{
    class Program
    {

        //name 프로세스 종료
        static void Close_process(string name)
        {
            try
            {
                string now = "VCC2 - Microsoft Visual Studio"; //꺼지면 안되는 프로그램
                //테스트를 위한 코드 : "아"라고 말하면 notePad(메모장)으로 인식하여 실행됨
                //if (name == "아")
                    name = "notePad";
                //====


                /*
                //특정 name 프로세스 종료
                Process[] processList = Process.GetProcessesByName(name);
                if(processList.Length>0)
                {
                    processList[0].CloseMainWindow();
                    Console.WriteLine("%s를 종료하였습니다.", name);
                }
                else
                {
                    Console.WriteLine("%s가 실행되어있지 않습니다.", name);
                }
                //======================
                */
                ///*
                /////현재 비주얼 스튜디오 제외 모든 프로세스 꺼보기
                if (name == "notePad")
                {
                    Process[] processList = Process.GetProcesses();//시스템의 모든 프로세스 정보 
                    Process rocessCurrent = Process.GetCurrentProcess();
                    foreach (Process p in processList)
                    {
                        Console.WriteLine("{0}를 종료하였습니다.\n", p.ProcessName);

                        if (p.ProcessName == now || p.ProcessName == "VCC2")
                             continue;
                       
                        p.CloseMainWindow();
                        
                    }
                    Console.WriteLine("프로세스를 종료 끝");
                }
                else
                {
                    Console.WriteLine("인식할 수 없는 명령어입니다.");
                }
                //*/
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }

        //컴퓨터 종료
        static void Computer_shutdown(string name)
        {
            if(name=="아")
                Process.Start("shutdown.exe", "-s -t 10");//10초 후 컴퓨터 종료
        }

        //컴퓨터 재부팅
        static void Computer_restart(string name)
        {
            if (name == "아")
                Process.Start("shutdown.exe", "-r -t 10");//10초 후 컴퓨터 재시작
        }

        //지정 경로(folder)에서 name 파일 찾기
        static void find(string name, string folder)//name="data"
        {
            try
            {
                string file = "*" + name + "*";//Notepad data.txt
                string dir = folder; //디렉토리
                //@"C:\Users\Bae hyunsu\Desktop" - 바탕화면

                DirectoryInfo Di = new DirectoryInfo(dir);
                FileInfo[] files = Di.GetFiles(file, SearchOption.AllDirectories);
                FileInfo correct = null;

                Console.WriteLine("현재 탐색 폴더 : {0}", folder);
                foreach (FileInfo s in files)
                {
                    if (Path.GetFileNameWithoutExtension(s.Name) == name)
                    {
                        correct = s;
                    }

                    //파일이름, 확장자, 풀 경로 출력
                    //Console.WriteLine("{0} , {1} , {2} ", Path.GetFileNameWithoutExtension(s.Name), Path.GetExtension(s.Name), Path.GetFullPath(s.Name));
                    Console.WriteLine("{0}", Path.GetFileName(s.Name));
                    //open_file(s.DirectoryName, s.Name); //해당 파일 열기
                }
                /*
                if (correct!=null)//정확한 파일명이 있으면 가장 나중에 출력
                    Console.WriteLine("정확 : {0}", Path.GetFullPath(correct.Name));
                */
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }

        //dir 폴더에 있는 file 열기
        static void open_file(string dir, string file)
        {
            try
            {
                System.Diagnostics.Process ps = new System.Diagnostics.Process();
                ps.StartInfo.FileName = file;
                ps.StartInfo.WorkingDirectory = dir;
                ps.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                ps.Start();
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally{ }
        }

    // [START speech_streaming_mic_recognize]
    static async Task<object> StreamingMicRecognizeAsync(int seconds)
        {

            string st = "";

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
                            LanguageCode = LanguageCodes.Korean.SouthKorea,//"en",
                        },
                        InterimResults = true,
                    }
                });
            // Print responses as they arrive.
            Task printResponses = Task.Run(async () => 
            {

                st = "";
                StringBuilder sb = new StringBuilder();
                while (await streamingCall.ResponseStream.MoveNext(
                    default(CancellationToken)))
                {
                    sb = new StringBuilder();
                    foreach (var result in streamingCall.ResponseStream
                        .Current.Results)
                    {
                        foreach (var alternative in result.Alternatives)
                        {
                            sb.Append(alternative.Transcript.ToString());                
                            Console.WriteLine(alternative.Transcript);
                        }
                    }
                }
                st = sb.ToString();
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

            /* 참고
             * string target= "http://www.microsoft.com";   string target = "ftp://ftp.microsoft.com";
             *string target = "C:\\Program Files\\Microsoft Visual Studio\\INSTALL.HTM";
            */
            string temp = "";
            string open_google = "http://google.com";
            string target ="";

            
            temp = st;

            Close_process(st);//2번 기능
            //Computer_shutdown(st);//5번 기능
            //Computer_restart(st);//6번 기능

            //find("data", @"C:\Users\Bae hyunsu\Desktop\졸업작품\default");//find(st);//2번 기능
            //open_file(@"C:\Users\Bae hyunsu\Desktop\졸업작품", "data.txt");//3번 기능
            /*
            if (st.Contains("구글 켜 줘"))
                target = open_google;
            if(target!="")
                System.Diagnostics.Process.Start(target);
            */
            return 0;
        }
        // [END speech_streaming_mic_recognize]
    

        public static int Main(string[] args)
        {
            return (int)StreamingMicRecognizeAsync(3).Result;
        }
    }
}
