using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAP2000v1;

namespace ProyectoPluginEjemplo
{
    static class Globales
    {
        public static cSapModel _SapModel;

        public static cPluginCallback _PluginCallback;

        public static VentanaFormulario _VentanaFormulario;
    }

    public class cPlugin : cPluginContract
    {
        public void Main(ref cSapModel SapModel, ref cPluginCallback ISapPlugin)
        {
            try
            {
                Globales._SapModel = SapModel;
                Globales._PluginCallback = ISapPlugin;


                Application.EnableVisualStyles();
                using (VentanaFormulario form = new VentanaFormulario())
                {
                    Globales._VentanaFormulario = form;

                    form.TopMost = true;
                    form.ShowDialog();

                }

                ISapPlugin.Finish(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public int Info(ref string Text)
        {
            try
            {
                Text = "";
                return 0;
            }
            catch (Exception)
            {
                return 1;
            }
        }
    }
    public class VentanaFormulario : Form
    {
        private Button button;

        public VentanaFormulario()
        {
            this.Text = "";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Height = 100;

            /*
            * Boton de Coeficientes de pandeo
            */
            button = new Button();
            button.Text = "";
            button.Top = 20;
            button.Left = 20;
            button.Width = 120;
            button.Click += Button_Click;

        }

        public void Button_Click(object sender, EventArgs e)
        {
            try
            {
                int ret = 0;

                Globales._VentanaFormulario.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
