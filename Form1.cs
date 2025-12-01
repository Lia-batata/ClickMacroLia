using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace LiaMacro
{
    public partial class Form1 : Form
    {
        private FlowLayoutPanel panelEtapas;
        private Button btnAdicionarEtapa;
        private Button btnIniciar;
        private Button btnParar;

        private bool macroRodando = false;
        private readonly InputSimulator inputSim = new InputSimulator();

        // hook mouse globals
        private bool aguardandoClique = false;
        private TextBox campoXAtual = null;
        private TextBox campoYAtual = null;
        private LowLevelMouseProc mouseProc;
        private IntPtr hookId = IntPtr.Zero;

        public Form1()
        {
            InitializeComponent();
            CriarInterface();
        }

        private void CriarInterface()
        {
            this.Text = "LiaMacro - Automação Inteligente";
            this.Size = new System.Drawing.Size(700, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;

            panelEtapas = new FlowLayoutPanel()
            {
                Location = new System.Drawing.Point(20, 70),
                Size = new System.Drawing.Size(640, 400),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10)
            };
            this.Controls.Add(panelEtapas);

            btnAdicionarEtapa = new Button()
            {
                Text = "+ Adicionar Etapa",
                Font = new Font("Segoe UI", 10),
                Size = new System.Drawing.Size(160, 35),
                Location = new System.Drawing.Point(20, 20)
            };
            btnAdicionarEtapa.Click += BtnAdicionarEtapa_Click;
            this.Controls.Add(btnAdicionarEtapa);

            btnIniciar = new Button()
            {
                Text = "▶ Iniciar",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new System.Drawing.Size(120, 40),
                Location = new System.Drawing.Point(200, 480)
            };
            btnIniciar.Click += BtnIniciar_Click;
            this.Controls.Add(btnIniciar);

            btnParar = new Button()
            {
                Text = "⏹ Parar",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new System.Drawing.Size(120, 40),
                Location = new System.Drawing.Point(340, 480)
            };
            btnParar.Click += BtnParar_Click;
            this.Controls.Add(btnParar);
        }


        // ADICIONAR ETAPA

        private void BtnAdicionarEtapa_Click(object sender, EventArgs e)
        {
            Panel painel = new Panel()
            {
                Width = 580,
                Height = 110,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(5)
            };
            panelEtapas.Controls.Add(painel);

            var lblEtapa = new Label()
            {
                Text = $"Etapa {panelEtapas.Controls.Count}",
                Location = new System.Drawing.Point(10, 10),
                AutoSize = true
            };
            painel.Controls.Add(lblEtapa);

            // miniatura e armazenamento de caminho (Tag)
            var picPreview = new PictureBox()
            {
                Location = new System.Drawing.Point(90, 10),
                Size = new System.Drawing.Size(60, 60),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                Tag = "" // caminho da imagem será guardado aqui
            };
            painel.Controls.Add(picPreview);

            var btnEscolherImagem = new Button()
            {
                Text = "Escolher Imagem",
                Location = new System.Drawing.Point(160, 10),
                Width = 120
            };
            btnEscolherImagem.Click += (s, ev) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Imagens|*.png;*.jpg;*.jpeg;*.bmp";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            // carrega imagem (clone para evitar arquivo bloqueado)
                            using (var img = Image.FromFile(ofd.FileName))
                            {
                                picPreview.Image = new Bitmap(img);
                            }
                            picPreview.Tag = ofd.FileName;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Erro ao carregar imagem: " + ex.Message);
                        }
                    }
                }
            };
            painel.Controls.Add(btnEscolherImagem);

            // Tempo (segundos)
            var lblTempo = new Label()
            {
                Text = "Tempo (s):",
                Location = new System.Drawing.Point(300, 15),
                AutoSize = true
            };
            painel.Controls.Add(lblTempo);

            var txtTempo = new TextBox()
            {
                Name = "txtTempo",
                Location = new System.Drawing.Point(370, 12),
                Width = 50,
                Text = "1"
            };
            painel.Controls.Add(txtTempo);

            // Coordenadas X/Y
            var lblX = new Label() { Text = "X:", Location = new System.Drawing.Point(150, 75), AutoSize = true };
            painel.Controls.Add(lblX);
            var txtX = new TextBox() { Name = "txtX", Location = new System.Drawing.Point(170, 72), Width = 60 };
            painel.Controls.Add(txtX);

            var lblY = new Label() { Text = "Y:", Location = new System.Drawing.Point(240, 75), AutoSize = true };
            painel.Controls.Add(lblY);
            var txtY = new TextBox() { Name = "txtY", Location = new System.Drawing.Point(260, 72), Width = 60 };
            painel.Controls.Add(txtY);

            // Botão selecionar posição
            var btnSelect = new Button()
            {
                Text = "Selecionar posição",
                Width = 140,
                Location = new System.Drawing.Point(330, 70)
            };
            btnSelect.Click += BtnSelecionarPosicao_Click;
            painel.Controls.Add(btnSelect);

            // Remover
            var btnRemover = new Button()
            {
                Text = "Remover",
                Location = new System.Drawing.Point(480, 10),
                Width = 80
            };
            btnRemover.Click += (s, ev) => panelEtapas.Controls.Remove(painel);
            painel.Controls.Add(btnRemover);
        }


        // CAPTURAR POSIÇÃO NA TELA (HOOK)

        private void BtnSelecionarPosicao_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            var painel = btn.Parent;

            campoXAtual = painel.Controls["txtX"] as TextBox;
            campoYAtual = painel.Controls["txtY"] as TextBox;

            if (campoXAtual == null || campoYAtual == null)
            {
                MessageBox.Show("Campos X/Y não encontrados nesta etapa.");
                return;
            }

            MessageBox.Show("Agora clique em qualquer ponto da tela. (Clique com o botão esquerdo)");

            aguardandoClique = true;

            // instala hook global
            mouseProc = MouseHookCallback;
            hookId = SetHook(mouseProc);
        }

        private IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && aguardandoClique && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                try
                {
                    MSLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                    int x = hookStruct.pt.x;
                    int y = hookStruct.pt.y;

                    // escreve nos campos e desinstala hook
                    this.Invoke((Action)(() =>
                    {
                        campoXAtual.Text = x.ToString();
                        campoYAtual.Text = y.ToString();
                        aguardandoClique = false;
                        MessageBox.Show($"Posição capturada: X={x}  Y={y}");
                    }));

                    UnhookWindowsHookEx(hookId);
                }
                catch
                {
                    // ignore
                }
            }

            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        // INICIAR MACRO

        private async void BtnIniciar_Click(object sender, EventArgs e)
        {

            if (macroRodando)
            {
                MessageBox.Show("O macro já está rodando!");
                return;
            }

            macroRodando = true;
            MessageBox.Show("Macro iniciada! O loop continuará até você clicar em PARAR.");

            // LOOP INFINITO
            while (macroRodando)
            {
                foreach (Panel p in panelEtapas.Controls)
                {
                    if (!macroRodando) break;

                    PictureBox picPreview = null;
                    TextBox txtTempo = null;
                    TextBox txtX = null;
                    TextBox txtY = null;

                    foreach (Control c in p.Controls)
                    {
                        if (c is PictureBox pic) picPreview = pic;
                        if (c.Name == "txtX") txtX = (TextBox)c;
                        if (c.Name == "txtY") txtY = (TextBox)c;

                        if (c is TextBox t &&
                            t.Name != "txtX" &&
                            t.Name != "txtY")
                        {
                            txtTempo = t;
                        }
                    }

                    string caminhoImagem = picPreview?.Tag as string;
                    bool temImagem = (!string.IsNullOrEmpty(caminhoImagem));
                    bool temCoord = (!string.IsNullOrEmpty(txtX?.Text) && !string.IsNullOrEmpty(txtY?.Text));

                    if (!temImagem && !temCoord)
                    {
                        MessageBox.Show("Nenhuma imagem OU coordenada definida nesta etapa.");
                        continue;
                    }

                    System.Drawing.Point clickPoint;

                    if (temImagem)
                    {
                        var pos = EncontrarImagemNaTela(caminhoImagem);

                        if (pos == null)
                        {
                            // NÃO para o macro > volta ao início do loop
                            continue;
                        }

                        clickPoint = pos.Value;
                    }
                    else
                    {
                        clickPoint = new System.Drawing.Point(
                            int.Parse(txtX.Text),
                            int.Parse(txtY.Text)
                        );
                    }

                    if (!double.TryParse(txtTempo?.Text, out double tempoSeg))
                        tempoSeg = 1;

                    int delay = (int)(tempoSeg * 1000);

                    // Movimento real do mouse
                    double normX = clickPoint.X * (65535.0 / (Screen.PrimaryScreen.Bounds.Width - 1));
                    double normY = clickPoint.Y * (65535.0 / (Screen.PrimaryScreen.Bounds.Height - 1));

                    inputSim.Mouse.MoveMouseTo(normX, normY);
                    inputSim.Mouse.LeftButtonClick();

                    // Aguarda o tempo configurado
                    await Task.Delay(delay);
                }

                // Final de todas as etapas > volta automaticamente ao início
            }

            MessageBox.Show("Macro finalizado!");
        }

     

        private void BtnParar_Click(object sender, EventArgs e)
        {
            macroRodando = false;
            MessageBox.Show("Macro parada.");
        }


        // FUNÇÃO OPEN CV — LOCALIZAR IMAGEM (ainda não esta funcionando como imaginei)

        private Mat CapturarTela()
        {
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            var bmp = new Bitmap(bounds.Width, bounds.Height);
            using (var g = Graphics.FromImage(bmp))
                g.CopyFromScreen(System.Drawing.Point.Empty, System.Drawing.Point.Empty, bounds.Size);
            return BitmapConverter.ToMat(bmp);
        }

        private System.Drawing.Point? EncontrarImagemNaTela(string caminho)
        {
            if (string.IsNullOrEmpty(caminho) || !File.Exists(caminho))
                return null;

            using (var imgTela = CapturarTela())
            using (var imgTemplate = new Mat(caminho, ImreadModes.Color))
            {
                if (imgTemplate.Empty())
                    return null;

                using (var resultado = new Mat())
                {
                    Cv2.MatchTemplate(imgTela, imgTemplate, resultado, TemplateMatchModes.CCoeffNormed);
                    Cv2.MinMaxLoc(resultado, out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

                    Console.WriteLine($"Similaridade: {maxVal:F2}");

                    if (maxVal >= 0.7)
                    {
                        return new System.Drawing.Point(
                            maxLoc.X + imgTemplate.Width / 2,
                            maxLoc.Y + imgTemplate.Height / 2
                        );
                    }
                }
            }

            return null;
        }


        // Windows mouse hook interop (necessário para captura de clique global)

        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x; public int y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}





