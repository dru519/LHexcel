using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace LH.NxSetup
{
    internal static class Program
    {
        internal static void OpenReset()
        {
            string directory=Path.Combine(Path.GetTempPath(),"LHExcelReset-"+Guid.NewGuid().ToString("N"));
            Package.RejectReparseAncestors(Path.GetTempPath());
            Directory.CreateDirectory(directory);
            string target=Path.Combine(directory,"내엑셀 설치 초기화.exe");
            byte[] bytes=File.ReadAllBytes(Assembly.GetExecutingAssembly().Location);
            using(var stream=new FileStream(target,FileMode.CreateNew,FileAccess.Write,FileShare.None))stream.Write(bytes,0,bytes.Length);
            if(Package.Hash(File.ReadAllBytes(target))!=Package.Hash(bytes))throw new IOException("초기화 실행 파일 확인 실패");
            Process.Start(new ProcessStartInfo(target,"--reset-ui"){UseShellExecute=true});
        }
        internal static string PreserveInstaller(string packageRoot)
        {
            byte[] binary = File.ReadAllBytes(Assembly.GetExecutingAssembly().Location);
            string target = Path.Combine(Path.GetDirectoryName(packageRoot), "Setup-" + Path.GetFileName(packageRoot) + ".exe");
            Package.RejectReparseAncestors(Path.GetDirectoryName(target));
            if (!File.Exists(target)) {
                using (var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    output.Write(binary, 0, binary.Length);
            }
            if ((File.GetAttributes(target) & FileAttributes.ReparsePoint) != 0 ||
                Package.Hash(File.ReadAllBytes(target)) != Package.Hash(binary))
                throw new IOException("Saved installer verification failed");
            return target;
        }
        static byte[] Resource(string name)
        {
            using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (input == null) throw new InvalidDataException("Missing installer resource: " + name);
                using (var output = new MemoryStream()) { input.CopyTo(output); return output.ToArray(); }
            }
        }

        [STAThread]
        static int Main(string[] args)
        {
            try
            {
                if(args.Length==1 && args[0]=="--reset-ui"){Reset.ResetWindow.Run();return 0;}
                byte[] zip = Resource("NxSetup.Package.zip");
                string version = Encoding.UTF8.GetString(Resource("NxSetup.Version")).Trim();
                string hash = Encoding.UTF8.GetString(Resource("NxSetup.Package.sha256")).Trim();
                if (Package.Hash(zip) != hash) throw new InvalidDataException("Embedded package is corrupt");
                var files = Package.Read(zip, version);
                if (args.Length == 1 && args[0] == "--verify-only") return 0;
                if (args.Length == 2 && args[1] == "--accept-current-user-changes" &&
                    (args[0] == "--install" || args[0] == "--uninstall" || args[0] == "--update" || args[0] == "--repair")) {
                    string root = args[0] != "--uninstall" ? Package.Stage(files, version, hash) :
                        Package.CachedPath(version, hash);
                    if (args[0] != "--uninstall") PreserveInstaller(root);
                    var installer = new NativeInstaller(version, hash, root, files, delegate(string message) {
                        File.AppendAllText(Path.Combine(Path.GetDirectoryName(root), "setup.log"),
                            DateTime.Now.ToString("s") + " " + message + Environment.NewLine, Encoding.UTF8);
                    });
                    if (args[0] == "--install") installer.Install();
                    else if(args[0] == "--update") installer.Update();
                    else if(args[0] == "--repair") installer.Repair();
                    else installer.RemoveInstalledVersion();
                    return 0;
                }
                bool interactiveUninstall=args.Length==1 && args[0]=="--uninstall";
                if (args.Length != 0 && !interactiveUninstall) throw new ArgumentException("Invalid installer arguments");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new SetupWindow(version, hash, files, interactiveUninstall));
                return 0;
            }
            catch (Exception error)
            {
                if (args.Length == 0 || (args.Length==1 && args[0]=="--uninstall")) MessageBox.Show(error.Message, "내엑셀 설치 중단", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else {
                    string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LHexcel", "Setup");
                    Directory.CreateDirectory(folder);
                    File.WriteAllText(Path.Combine(folder, "last-error.txt"), error.ToString(), Encoding.UTF8);
                }
                return 1;
            }
        }
    }

    internal sealed class SetupWindow : Form
    {
        readonly string version, hash;
        readonly Dictionary<string, byte[]> files;
        readonly TextBox output;
        readonly FlowLayoutPanel buttons;
        readonly FlowLayoutPanel utilities;
        readonly Label heading;
        readonly RadioButton installButton, repairButton, removeButton, checkButton;
        readonly Panel page;
        readonly TextBox summary;
        readonly ProgressBar progress;
        readonly Button launchExcel;
        readonly Button backButton, nextButton, closeButton, detailsButton;
        string pendingAction;
        string screen="welcome", installedVersion;
        bool completed;
        bool busy;

        public SetupWindow(string version, string hash, Dictionary<string, byte[]> files, bool uninstallOnOpen=false)
        {
            this.version = version; this.hash = hash; this.files = files;
            Text = "내엑셀 " + version + " 설치";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("NxSetup.Brand.Icon"))
                if (stream != null) using (var sourceIcon = new Icon(stream)) Icon = (Icon)sourceIcon.Clone();
            Font = new Font("맑은 고딕", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(680, 480);
            MinimumSize = new Size(680, 480);
            StartPosition = FormStartPosition.CenterScreen;
            heading = new Label { Dock = DockStyle.Top, Height = 85, Padding = new Padding(82,18,18,18),
                BackColor=Color.FromArgb(40,100,103),ForeColor=Color.White,Font=new Font(Font.FontFamily,12F,FontStyle.Bold),
                Text = "내엑셀 설치를 시작합니다\r\n" + version };
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("NxSetup.Brand.Mark"))
                if (stream != null) using (var image = Image.FromStream(stream)) {
                    var mark = new PictureBox {Image=new Bitmap(image),Size=new Size(48,48),Location=new Point(18,18),SizeMode=PictureBoxSizeMode.Zoom};
                    mark.Disposed += delegate {mark.Image.Dispose();};
                    heading.Controls.Add(mark);
                }
            output = new TextBox { Dock = DockStyle.Bottom, Height=130, Visible=false, Multiline = true, ReadOnly = true,
                ScrollBars = ScrollBars.Vertical, WordWrap = true };
            buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 175, Padding = new Padding(18,4,18,4),
                FlowDirection = FlowDirection.TopDown,WrapContents=false,Visible=false };
            installButton=AddChoice("설치", "update");
            repairButton=AddChoice("복구 — 설치 파일과 추가 기능 등록 복원", "repair");
            removeButton=AddChoice("제거 — 내엑셀 자동 실행과 추가 기능 등록 제거", "uninstall");
            checkButton=AddChoice("설치 상태 확인 — 파일과 등록 정보 검사", "check");
            page=new Panel {Dock=DockStyle.Fill,Padding=new Padding(18)};
            summary=new TextBox {Dock=DockStyle.Fill,Multiline=true,ReadOnly=true,BorderStyle=BorderStyle.None,
                BackColor=SystemColors.Control,ScrollBars=ScrollBars.Vertical,TabStop=false};
            progress=new ProgressBar {Dock=DockStyle.Bottom,Height=20,Style=ProgressBarStyle.Marquee,Visible=false};
            launchExcel=new Button {Text="Excel 실행",AutoSize=true,Visible=false};
            launchExcel.Click+=delegate {
                try {OfferExcelLaunch(pendingAction,()=>true,LaunchExcel);launchExcel.Enabled=false;}
                catch(Exception error){Log(error.ToString());summary.Text="작업은 완료되었지만 Excel을 시작하지 못했습니다. Excel을 직접 실행하거나 다시 시도해 주세요.\r\n"+error.Message;}
            };
            page.Controls.Add(summary);page.Controls.Add(progress);
            var navigation=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=52,Padding=new Padding(10),FlowDirection=FlowDirection.RightToLeft};
            closeButton=new Button {Text="취소",AutoSize=true};closeButton.Click+=delegate {Close();};
            nextButton=new Button {Text="다음 >",AutoSize=true,Visible=true};
            nextButton.Click+=delegate {
                if(screen=="welcome") {if(installedVersion==null)ReviewAction("update");else ShowChoices();return;}
                if(screen=="choices") {ReviewAction(pendingAction);return;}
                if(completed) {
                    if(pendingAction=="check"){ShowChoices();return;}
                    Close();
                } else RunAction(pendingAction);
            };
            backButton=new Button {Text="< 뒤로",AutoSize=true,Visible=false};backButton.Click+=delegate {
                if(screen=="choices" || installedVersion==null)ShowWelcome();else ShowChoices();
            };
            detailsButton=new Button {Text="상세 로그 보기",AutoSize=true};detailsButton.Click+=delegate {output.Visible=!output.Visible;detailsButton.Text=output.Visible ? "상세 로그 숨기기" : "상세 로그 보기";};
            navigation.Controls.Add(closeButton);navigation.Controls.Add(nextButton);navigation.Controls.Add(launchExcel);navigation.Controls.Add(backButton);navigation.Controls.Add(detailsButton);
            Controls.Add(page);Controls.Add(output);Controls.Add(buttons);Controls.Add(heading);Controls.Add(navigation);
            CancelButton=closeButton;
            string runtime = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LHexcel", "NxHost", version.Split('_')[1]);
            output.Text = "설치 위치: " + runtime + "\r\n" +
                "자동 실행: " + Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel", "XLSTART", "내엑셀 " + version + ".xlam") + "\r\n" +
                "\r\n설치 후 Excel을 열면 내엑셀 탭이 자동으로 표시됩니다.\r\n"+
                "사용자 설정은 업데이트·제거 시에도 보관합니다.";
            summary.Text="이 프로그램은 현재 Windows 사용자에게 내엑셀을 설치합니다.\r\n\r\n계속하려면 ‘다음’을 누르세요. 설치는 다음 화면에서 ‘설치’를 눌러야 시작됩니다.\r\n\r\nExcel 문서를 저장하고 Excel을 모두 닫아 주세요.\r\n사용자 설정은 업데이트·제거 시에도 보관합니다.";
            var folder = new Button { Text = "설치 폴더", AutoSize = true };
            folder.Click += delegate {
                if(Directory.Exists(runtime)) Process.Start("explorer.exe", "\"" + runtime + "\"");
                else MessageBox.Show("설치 후 폴더를 열 수 있습니다.", Text);
            };
            utilities=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=40,Padding=new Padding(10,0,10,0)};
            utilities.Controls.Add(folder);
            var history=new Button {Text="이전 설치 기록 정리",AutoSize=true};
            history.Click+=delegate { ShowHistory(); };utilities.Controls.Add(history);
            var reset=new Button {Text="완전 삭제 (설치 이력 포함)",AutoSize=true};
            reset.Click+=delegate {
                if(busy)return;
                try{Program.OpenReset();Close();}
                catch(Exception error){MessageBox.Show(this,error.Message,"완전 삭제 준비 오류");}
            };
            utilities.Controls.Add(reset);
            Controls.Add(utilities);utilities.BringToFront();
            FormClosing += delegate(object sender, FormClosingEventArgs e) {
                if (busy) { e.Cancel = true; MessageBox.Show("진행 중인 설치·복원이 끝난 뒤 닫아 주세요."); }
            };
            Shown += delegate {
                if(files.Count>0) RefreshActions();
                if(uninstallOnOpen && nextButton.Enabled) BeginInvoke((Action)delegate {
                    if(removeButton.Enabled)ReviewAction("uninstall");
                    else ShowAlreadyRemoved();
                });
            };
            AcceptButton=nextButton;
        }

        internal static bool OfferExcelLaunch(string action, Func<bool> confirm, Action launch)
        {
            if(action!="update" || !confirm()) return false;
            launch();
            return true;
        }

        void ShowHistory()
        {
            try {
                var installer=new NativeInstaller(version,hash,Package.CachedPath(version,hash),files,Log);
                string[] versions=installer.RemovedHistory();
                using(var dialog=new Form {Text="이전 설치 기록 정리",ClientSize=new Size(520,380),MinimumSize=new Size(520,380),AutoScaleMode=AutoScaleMode.Dpi,Font=Font,StartPosition=FormStartPosition.CenterParent,MinimizeBox=false,MaximizeBox=false}) {
                    var note=new Label {Dock=DockStyle.Top,Height=78,Text=versions.Length==0 ? "정리할 제거 완료 기록이 없습니다.\r\n현재 설치와 제거가 확인되지 않은 기록은 보존합니다." : "정리할 이전 버전을 선택하세요. 현재 설치와 사용자 설정은 유지합니다.\r\n선택한 기록은 복구 보관 폴더로 이동합니다. 디스크 공간은 줄어들지 않습니다.",Padding=new Padding(10)};
                    var list=new CheckedListBox {Dock=DockStyle.Fill,CheckOnClick=true};list.Items.AddRange(versions);
                    var selection=HistorySelectionTools(list);
                    var cache=new CheckBox {Dock=DockStyle.Bottom,Height=32,Text="선택한 버전의 설치 파일·캐시도 함께 정리",Checked=false};
                    var actions=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=45};
                    var apply=new Button {Text="선택 기록 정리",AutoSize=true,Enabled=versions.Length>0};var cancel=new Button {Text="닫기",DialogResult=DialogResult.Cancel};
                    cache.Enabled=versions.Length>0;
                    actions.Controls.Add(apply);actions.Controls.Add(cancel);
                    apply.Click+=delegate {
                        var chosen=new List<string>();foreach(object item in list.CheckedItems)chosen.Add((string)item);
                        if(chosen.Count==0){MessageBox.Show(dialog,"정리할 버전을 선택하세요.");return;}
                        if(MessageBox.Show(dialog,String.Join(", ",chosen.ToArray())+"\r\n선택한 기록을 복구 보관 폴더로 이동할까요?","기록 정리 확인",MessageBoxButtons.OKCancel)!=DialogResult.OK)return;
                        try{string location=installer.ArchiveHistory(chosen.ToArray(),cache.Checked);MessageBox.Show(dialog,"정리했습니다. 복구 보관 위치:\r\n"+location);dialog.Close();}
                        catch(Exception error){MessageBox.Show(dialog,error.Message,"기록 정리 중단");}
                    };
                    dialog.Controls.Add(list);dialog.Controls.Add(selection);dialog.Controls.Add(cache);dialog.Controls.Add(actions);dialog.Controls.Add(note);dialog.CancelButton=cancel;dialog.ShowDialog(this);
                }
            } catch(Exception error){MessageBox.Show(this,error.Message,"설치 기록 확인");}
        }

        internal static FlowLayoutPanel HistorySelectionTools(CheckedListBox list)
        {
            var panel=new FlowLayoutPanel {Dock=DockStyle.Top,Height=36};
            var all=new Button {Text="전체 선택",AutoSize=true,Enabled=list.Items.Count>0};
            var none=new Button {Text="전체 해제",AutoSize=true,Enabled=list.Items.Count>0};
            var count=new Label {AutoSize=true,Margin=new Padding(12,6,0,0)};
            Action refresh=()=>count.Text=list.CheckedItems.Count+" / "+list.Items.Count+"개 선택";
            all.Click+=delegate {for(int i=0;i<list.Items.Count;i++)list.SetItemChecked(i,true);refresh();};
            none.Click+=delegate {for(int i=0;i<list.Items.Count;i++)list.SetItemChecked(i,false);refresh();};
            list.ItemCheck+=delegate(object sender,ItemCheckEventArgs e) {
                int total=list.CheckedItems.Count+(e.NewValue==CheckState.Checked ? 1 : 0)-(e.CurrentValue==CheckState.Checked ? 1 : 0);
                count.Text=total+" / "+list.Items.Count+"개 선택";
            };
            panel.Controls.Add(all);panel.Controls.Add(none);panel.Controls.Add(count);refresh();return panel;
        }

        static void LaunchExcel()
        {
            string path=NativeInstaller.FindExcel();
            ushort machine=NativeInstaller.Machine(path);
            if(machine!=0x14c && machine!=0x8664) throw new InvalidOperationException("지원하는 Excel 실행 파일을 찾을 수 없습니다.");
            using(var process=Process.Start(new ProcessStartInfo(path) { UseShellExecute=false, WindowStyle=ProcessWindowStyle.Normal })) {
                if(process==null) throw new IOException("Excel을 시작하지 못했습니다.");
            }
        }

        RadioButton AddChoice(string label, string action)
        {
            var button = new RadioButton { Text = label, AutoSize = true,Margin=new Padding(0,5,0,5) };
            button.CheckedChanged += delegate {if(button.Checked)pendingAction=action;};
            buttons.Controls.Add(button);
            return button;
        }

        void ShowChoices()
        {
            screen="choices";completed=false;buttons.Visible=true;backButton.Visible=true;nextButton.Visible=true;nextButton.Text="다음 >";
            launchExcel.Visible=false;closeButton.Visible=true;closeButton.Text="취소";AcceptButton=nextButton;
            summary.Text="원하는 작업을 선택한 뒤 ‘다음’을 누르세요.\r\n사용자 설정은 업데이트·제거 시에도 보관합니다.";
            RefreshActions();
            heading.Text="내엑셀 설치 관리\r\n현재 설치: "+(installedVersion ?? "없음");
            if(installButton.Visible && installButton.Enabled){installButton.Checked=true;pendingAction="update";}
            else if(repairButton.Enabled){repairButton.Checked=true;pendingAction="repair";}
            else {checkButton.Checked=true;pendingAction="check";}
        }

        void ShowAlreadyRemoved()
        {
            screen="done";completed=true;pendingAction="uninstall";
            buttons.Visible=false;backButton.Visible=false;closeButton.Visible=false;nextButton.Visible=true;nextButton.Text="마침";
            heading.Text="제거할 내엑셀이 없습니다";
            summary.Text="현재 활성 내엑셀 설치가 확인되지 않았습니다.\r\n이미 제거된 상태라면 추가 제거 작업이 필요하지 않습니다.\r\n보관된 사용자 설정과 과거 기록은 그대로 유지됩니다.";
        }

        void ShowWelcome()
        {
            screen="welcome";completed=false;buttons.Visible=false;backButton.Visible=false;nextButton.Visible=true;
            nextButton.Text="다음 >";closeButton.Visible=true;closeButton.Text="취소";launchExcel.Visible=false;
            heading.Text="내엑셀 설치를 시작합니다\r\n"+version;
            summary.Text="설치·업데이트·복구·제거를 진행할 수 있습니다.\r\n계속하려면 ‘다음’을 누르세요.\r\n\r\nExcel 문서를 저장하고 모두 닫아 주세요.\r\n사용자 설정은 보관합니다.";
            RefreshActions();
        }

        void ReviewAction(string action)
        {
            if(busy)return;
            screen="review";pendingAction=action;completed=false;buttons.Visible=false;backButton.Visible=true;nextButton.Visible=true;
            nextButton.Text=action=="uninstall" ? "제거" : action=="repair" ? "복구" : action=="check" ? "확인 시작" : installedVersion==null ? "설치" : "업데이트";
            heading.Text=nextButton.Text+" 준비\r\n아래 내용을 확인한 뒤 진행하세요.";
            summary.Text=(action=="uninstall" ? "설치된 내엑셀의 자동 실행 파일과 추가 기능 등록을 제거합니다.\r\n\r\n제거 대상: 내엑셀 "+(installedVersion ?? version)+"\r\n"+Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LHexcel","NxHost",(installedVersion ?? version).Split('_')[1])+"\r\n\r\n사용자 설정과 설치 기록은 보관합니다.\r\n‘제거’를 누르면 시작합니다." :
                action=="check" ? "설치 파일과 추가 기능 등록 상태를 검사합니다." :
                    "내엑셀 "+version+" · 현재 사용자에게 "+nextButton.Text+"합니다.\r\n\r\n설치 위치 (자동 지정):\r\n"+Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LHexcel","NxHost",version.Split('_')[1])+"\r\n\r\n설치 후 Excel을 열면 내엑셀 탭이 자동으로 표시됩니다.")+"\r\n\r\nExcel 문서를 저장하고 모든 Excel 창을 닫아 주세요.";
            AcceptButton=nextButton;
        }

        void RefreshActions()
        {
            try {
                var installer=new NativeInstaller(version,hash,Package.CachedPath(version,hash),files,delegate{});
                string installed=installer.InstalledVersion();
                installedVersion=installed;
                bool refresh=installer.HasPackageRefresh();
                installButton.Visible=installed!=version || refresh;
                installButton.Text=installed==null ? "설치 — 내엑셀 추가 기능 설치" : "업데이트 — 이 패키지로 설치 갱신";
                repairButton.Enabled=installed==version && !refresh;
                removeButton.Enabled=installed!=null;
                nextButton.Enabled=true;
            } catch(Exception error) {
                installButton.Enabled=false;repairButton.Enabled=false;removeButton.Enabled=false;
                nextButton.Enabled=false;
                summary.Text="설치 상태를 확인할 수 없습니다.\r\n"+error.Message+"\r\n"+
                    NativeInstaller.RegistrationSummary()+"\r\n‘완전 삭제 (설치 이력 포함)’에서 초기화한 뒤 이 설치 파일을 다시 실행하세요.";
                output.AppendText("\r\n설치 상태 확인 필요: "+error.Message);
            }
        }

        void Log(string text)
        {
            try {
                string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "LHexcel", "Setup", version);
                Package.RejectReparseAncestors(directory);
                Directory.CreateDirectory(directory);
                File.AppendAllText(Path.Combine(directory, "setup-ui.log"),
                    DateTime.Now.ToString("s") + " " + text + Environment.NewLine, Encoding.UTF8);
            } catch (IOException) { /* The visible progress remains available if diagnostic logging fails. */ }
              catch (UnauthorizedAccessException) { }
            BeginInvoke((Action)delegate {
                if (output.TextLength > 60000) output.Clear();
                output.AppendText("\r\n" + text);
            });
        }

        void RunAction(string action)
        {
            if (busy) return;
            busy = true; buttons.Enabled = false;utilities.Enabled=false;
            screen="progress";
            nextButton.Enabled=false;backButton.Enabled=false;closeButton.Enabled=false;progress.Visible=true;
            summary.Text="작업 중입니다. 설치 파일과 등록 정보를 확인하고 있습니다.\r\n완료될 때까지 창을 닫지 마세요.";
            heading.Text=(action=="uninstall" ? "제거 중" : action=="check" ? "상태 확인 중" : action=="repair" ? "복구 중" : "설치 중")+"\r\n작업이 끝날 때까지 잠시 기다려 주세요.";
            var worker = new BackgroundWorker();
            string checkResult="";
            worker.DoWork += delegate {
                string root = action == "uninstall" || action=="check" ? Package.CachedPath(version, hash) : Package.Stage(files, version, hash);
                if (action != "uninstall" && action!="check") {
                    Log("제거할 때 다시 실행할 설치 파일: " + Program.PreserveInstaller(root));
                    Log("검증된 설치 패키지 보존 위치: " + root);
                    Package.VerifyDirectory(root, files);
                }
                var installer = new NativeInstaller(version, hash, root, files, Log);
                if (action == "check") {
                    string status=installer.Check()+"\r\n"+installer.MaintenanceStatus();
                    Log(status);
                    // RunWorkerCompleted displays the result without closing the wizard.
                    checkResult=status;
                } else {
                    if (action == "update") installer.Update();
                    else if(action == "repair") installer.Repair();
                    else installer.RemoveInstalledVersion();
                }
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e) {
                busy = false; buttons.Enabled = true;utilities.Enabled=true; worker.Dispose();
                nextButton.Enabled=true;backButton.Enabled=true;closeButton.Enabled=true;progress.Visible=false;
                if (e.Error != null) {
                    screen="error";
                    Log("실패 작업: "+action+" · 패키지: "+hash+"\r\n"+e.Error.ToString());
                    output.AppendText("\r\n중단: " + e.Error.Message);
                    heading.Text="작업을 완료하지 못했습니다";
                    summary.Text=e.Error.Message+"\r\n\r\n"+NativeInstaller.RegistrationSummary()+
                        "\r\n\r\n설치 기록과 등록 경로가 다르면 ‘완전 삭제 (설치 이력 포함)’에서 초기화한 뒤 설치 파일을 다시 실행하세요.";
                    nextButton.Text="다시 시도";closeButton.Text="닫기";
                    return;
                }
                heading.Text=(action=="uninstall" ? "제거 완료" : action=="check" ? "확인 완료" : action=="repair" ? "복구 완료" : "설치 완료");
                screen="done";completed=true;backButton.Visible=false;nextButton.Text=action=="check" ? "설치 관리로 돌아가기" : "마침";closeButton.Visible=action=="check";closeButton.Text="닫기";
                summary.Text=action=="uninstall" ? "내엑셀 제거가 완료되었습니다.\r\n사용자 설정과 설치 기록은 보관됩니다.\r\n‘마침’을 누르면 제거 프로그램이 종료됩니다." : action=="check" ? checkResult+"\r\n\r\n‘설치 관리로 돌아가기’를 눌러 다음 작업을 선택하세요." : "내엑셀 "+version+"의 작업을 완료했습니다.\r\n다음 Excel 실행부터 내엑셀을 사용할 수 있습니다.\r\n‘마침’을 누르면 설치 프로그램이 종료됩니다.";
                launchExcel.Visible=action=="update";
            };
            worker.RunWorkerAsync();
        }

    }
}
