using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using NAudio;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.Speech.V1;
using Grpc.Auth;
using System.IO;
using System.Threading;
// Install-Package Google.Cloud.Speech.V1 -Version 1.0.0-beta08 -Pre


//========= 17.07.24 ============
//===== 작성자 : 배현수 =========
using System.Diagnostics;
//===============================

namespace VCC_wpf
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        static string program = "";
        static string command = "";
        static string voice = "";
        static SortedList<string, string> userCommand = new SortedList<string, string>();
        static List<string> listOfuserCommand = new List<string>();
        int sizeOfuserCommand = 0;
        [DllImport("user32.dll")]
        public static extern void keybd_event(uint vk, uint scan, uint flags, uint extraInfo);

        /*=========== 17.07.24 ============
         ======== 작성자 : 배현수 =========
         1. Close_process(string name) : name 프로세스 종료
         2. Computer_shutdown() : 컴퓨터 종료
         3. Computer_restart() : 컴퓨터 재부팅
         4. find(string name, string folder) : 지정 경로(folder)에서 name 파일 찾기
         5. open_file(string dir, string file) : dir 폴더에 있는 file 열기
         ==================================*/
        static string DefaultFolder = @"C:\Users\Bae hyunsu\Desktop";//바탕화면


        //1. name 프로세스 종료
        static void Close_process(string name)
        {
            try
            {
                //string now = "VCC2 - Microsoft Visual Studio"; //꺼지면 안되는 프로그램
                //테스트를 위한 코드 : "아"라고 말하면 notePad(메모장)으로 인식하여 실행됨
                //if (name == "아")
                //name = "notePad";
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

                //Alt + Ctrl + F4
                //현재 비주얼 스튜디오 제외 모든 프로세스 꺼보기
                string now = "devenv";
                string now1 = "VCC_wpf";
                if (name == "")
                {
                    Process[] processList = Process.GetProcesses();//시스템의 모든 프로세스 정보 
                    Process rocessCurrent = Process.GetCurrentProcess();
                    foreach (Process p in processList)
                    {
                        if (p.ProcessName.Equals(now)||p.ProcessName.Equals(now1)) continue;//현재 프로그램 제외
                        Console.WriteLine("{0}를 종료하였습니다.\n", p.ProcessName);
                        //p.CloseMainWindow();
                        p.Kill();
                    }
                    Console.WriteLine("프로세스를 종료 끝");
                }
                else
                {
                    Console.WriteLine("인식할 수 없는 명령어입니다.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }

        //2. 컴퓨터 종료
        static void Computer_shutdown()
        {
            //if (name == "아")
                Process.Start("shutdown.exe", "-s -t 10");//10초 후 컴퓨터 종료
        }

        //3. 컴퓨터 재부팅
        static void Computer_restart()
        {
            //if (name == "아")
                Process.Start("shutdown.exe", "-r -t 10");//10초 후 컴퓨터 재시작
        }

        //4. 지정 경로(folder)에서 name 파일 찾기
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
                    if (System.IO.Path.GetFileNameWithoutExtension(s.Name) == name)
                    {
                        correct = s;

                        //일단 정확한 경우만 출력하게 만들어 놓음
                        open_file(dir, s.Name);
                    }

                    //파일이름, 확장자, 풀 경로 출력
                    //Console.WriteLine("{0} , {1} , {2} ", Path.GetFileNameWithoutExtension(s.Name), Path.GetExtension(s.Name), Path.GetFullPath(s.Name));
                    Console.WriteLine("{0}", System.IO.Path.GetFileName(s.Name));
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

        //5. dir 폴더에 있는 file 열기
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
            finally { }
        }
        //=================================
        //=================================

        /*=========== 17.07.24 ============
        ======== 작성자 : 배현수 =========
        1. Close_process(string name) : name 프로세스 종료
        2. Computer_shutdown() : 컴퓨터 종료
        ==================================*/


        //DefaultFolder = @"C:\Users\Bae hyunsu\Desktop" 변경 
        static void Change_DefaultFolder(string dir)
        {
            try
            {
                DefaultFolder = dir;
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }
        //==================================================
        //==================================================

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
        static void SearchWeb(string searchword, string web = "")
        {
            string query = MakeQuery(searchword);
            string url = null;
            if (web.Equals("네이버"))
            {
                url = "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=" + query;
            }
            else
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

            if (command.Equals("open") && !program.Equals(""))        // 프로그램 실행
            {
                //process.Start("control");
                //// CMD에 보낼 명령어를 입력.
                //process.StandardInput.Write();
                //process.StandardInput.Close();
                /* 탐색기 기능 인수 구분 구현 예정*/
                if (program.Equals("internet"))
                    System.Diagnostics.Process.Start("explorer", "http://google.com");
                else
                    System.Diagnostics.Process.Start(program);
            }
            else if (command.Equals("empty"))                      // 휴지통 비우기
            {
                /* 휴지통 비우기 확인(Y/N) 창 띄우기 예정 */
                System.Diagnostics.Process.Start("cmd.exe", "/c rd /s /q %systemdrive%\\$Recycle.bin");
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
            else if (voice.Contains("휴지통") && (voice.Contains("비워") || voice.Contains("비우")))
            {
                command = "empty";
            }
            else if (voice.Contains("프로그램") && voice.Contains("등록"))
            {

            }
            else if (voice.Contains("스크린샷"))
            {
                /* VK_LWIN = 0x5B , VK_SNAPSHOT = 0x2C
                 * WIN + PRINT_SCREEN = 전체 캡쳐 후 사진 - 스크린샷 폴더에 저장
                 * flags : 0x00 (DOWN) 0x02 (UP) */
                keybd_event(0x5B, 0, 0x00, 0);
                keybd_event(0x2C, 0, 0x00, 0);
                keybd_event(0x5B, 0, 0x02, 0);
                keybd_event(0x2C, 0, 0x02, 0);
            }
            else if (voice.Contains("검색"))
            {
                command = "search";
                // 검색어 추출
                string searchword = voice.Substring(0, voice.IndexOf("검색"));
                if (searchword == null)
                {
                    Console.WriteLine("검색할 검색어가 없습니다.");
                    return;
                }
                // 공백 제거
                searchword = searchword.Trim();
                SearchWeb(searchword);
            }


            /*=========== 17.07.24 ============
             ======== 작성자 : 배현수 ========= */

            //1.Close_process(string name) : name 프로세스 종료
            else if (voice.Contains("모든 창 꺼 줘"))
            {
                //command = "close";
                // 종료 프로세스 이름 추출
                /*string ProcessName = voice.Substring(0, voice.IndexOf("종료"));
                if (ProcessName == null)
                {
                    Console.WriteLine("종료할 프로세스를 찾을 수 없습니다.");
                    return;
                }
                
                // 공백 제거
                ProcessName = ProcessName.Trim();
                */
                //Close_process(ProcessName);
                Close_process("");
            }
            
            //2. Computer_shutdown() : 컴퓨터 종료
            else if (voice.Contains("컴퓨터 꺼줘"))
            {
                Computer_shutdown();
            }
            
            //3. Computer_restart() : 컴퓨터 재부팅
            else if (voice.Contains("다시 시작"))
            {
                Computer_restart();
            }

            //4. find(string name, string folder) : 지정 경로(folder)에서 name 파일 찾기
            else if (voice.Contains("켜줘"))
            {
                //command = "open";//명령어 겹침..
                // 파일 이름 추출
                string Name = voice.Substring(0, voice.IndexOf("켜줘"));
                if (Name == null)
                {
                    Console.WriteLine("올바른 프로그램을 찾을 수 없습니다.");
                    return;
                }
                // 공백 제거
                Name = Name.Trim();
                find(Name, DefaultFolder);//디폴트 폴더에서 Name 파일 찾아서 실행
            }
            //==================================
            //==================================

            else
            {
                foreach (string name in listOfuserCommand)
                {
                    if (voice.Contains(name))
                        program = userCommand[name];
                }
            }
            if (voice.Contains("실행") || voice.Contains("켜"))
                command = "open";
        }
        // [START speech_streaming_mic_recognize]
        static async Task<int> StreamingMicRecognizeAsync(int seconds)
        {
            if (NAudio.Wave.WaveIn.DeviceCount < 1)
            {
                //Console.WriteLine("No microphone!");
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
                while (await streamingCall.ResponseStream.MoveNext(
                    default(CancellationToken)))
                {
                    foreach (var result in streamingCall.ResponseStream
                        .Current.Results)
                    {
                        foreach (var alternative in result.Alternatives)
                        {
                            //Console.WriteLine(alternative.Transcript);
                            //MessageBox.Show(alternative.Transcript);
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
            //Console.WriteLine("Speak now.");
            MessageBox.Show("Speak now");
            await Task.Delay(TimeSpan.FromSeconds(seconds));
            // Stop recording and shut down.
            waveIn.StopRecording();
            lock (writeLock) writeMore = false;
            await streamingCall.WriteCompleteAsync();
            await printResponses;
            /* 명령어 체크 후 실행 */
            MessageBox.Show(voice);
            CheckCommand(voice);
            ExecCommand();
            return 0;
        }
        // [END speech_streaming_mic_recognize]
       
        // [END speech_streaming_mic_recognize]
        private void textBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in file)
                {
                    this.textBox_filePath.Text = str;
                }
            }
        }

        private void textBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void textBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void button_register_Click(object sender, RoutedEventArgs e)
        {
            string program_name = textBox_registerName.Text;
            string program_filePath = "\"" + textBox_filePath.Text + "\"";
            /* 등록하려는 명령어가 한글인지 검사 */
            char[] charArr = program_name.ToCharArray();
            foreach (char c in charArr)
            {
                if (char.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.OtherLetter)
                {
                    //Console.WriteLine("한글이 아닌 문자가 있습니다. 한글로 프로그램 이름을 등록해주십시오.");
                    MessageBox.Show("한글이 아닌 문자가 있습니다. 한글로 프로그램 이름을 등록해주십시오.");
                    return;
                }
            }
            ///* 등록하려는 명령어가 이미 있는 명령어인지 검사 (구현 예정) */
            //if (userCommand.IndexOfKey(program_name) < 0)
            //    MessageBox.Show("이미 등록되어 있는 명령어입니다.");
            ///* 등록하려는 프로그램이 이미 있는 프로그램인지 검사 (구현 예정) */
            //if (userCommand.IndexOfValue(program_filePath) < 0)
            //    MessageBox.Show("이미 등록되어 있는 프로그램입니다.");
            ///* 사용자의 명령어 및 프로그램 등록 */
            //else
            //{
                userCommand.Add(program_name, program_filePath);
                listOfuserCommand.Add(program_name);
                MessageBox.Show("등록되었습니다.");
           // }
            
        }

        private async void button_voice_Click(object sender, RoutedEventArgs e)
        {
            int getResult = await StreamingMicRecognizeAsync(5);
            //int resultOfMicRecognize = await getResult;
            if (getResult == -1)
                MessageBox.Show("연결된 마이크 기기가 없습니다.");
            return;
        }
    }
}
