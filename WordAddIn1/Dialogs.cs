using System.Windows.Forms;


namespace WordAddIn1
{
    public partial class Ribbon1
    {
        public string NameInputDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 300,
                Height = 140,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 25, Top = 15, Text = text };
            TextBox textBox = new TextBox() { Left = 25, Top = 40, Width = 225 };
            Button confirmation = new Button() { Text = caption, Left = 175, Width = 75, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        public string ModelNameTakenDialog(string TakenModelName)
        {
            Form prompt = new Form()
            {
                Width = 235,
                Height = 140,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 25, Top = 15, Width = 200, Text = "This model name is already taken." };
            Label textLabel2 = new Label() { Left = 25, Top = 40, Width = 200, Text = "Do you want to overwrite it?" };
            Button NOTconfirmation = new Button() { Text = "NO", Left = 25, Width = 75, Top = 65, DialogResult = DialogResult.OK };
            Button confirmation = new Button() { Text = "YES", Left = 110, Width = 75, Top = 65, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { TakenModelName += "-ToOverwrite"; prompt.Close(); Globals.Ribbons.Ribbon1.Overwrite = true; };
            NOTconfirmation.Click += (sender, e) => { TakenModelName = NameInputDialog("Model name:", "TRAIN!"); Globals.Ribbons.Ribbon1.Overwrite = false; prompt.Close(); };
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(NOTconfirmation);
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textLabel2);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? TakenModelName : "";
        }

        public string ProjectNameTakenDialog(string TakenModelName)
        {
            Form prompt = new Form()
            {
                Width = 235,
                Height = 115,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 25, Top = 15, Width = 200, Text = "This project name is already taken." };
            Button NOTconfirmation = new Button() { Text = "OK", Left = 70, Width = 70, Top = 40, DialogResult = DialogResult.OK };
            NOTconfirmation.Click += (sender, e) => { TakenModelName = NameInputDialog("New Project name:", "CREATE!"); prompt.Close(); };
            prompt.Controls.Add(NOTconfirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = NOTconfirmation;

            return prompt.ShowDialog() == DialogResult.OK ? TakenModelName : "";
        }

        public string TextMessageOkDialog(string SomeString = "")
        {
            Form prompt = new Form()
            {
                Width = 235,
                Height = 115,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 25, Top = 15, Width = 200, Text = SomeString };
            Button NOTconfirmation = new Button() { Text = "OK", Left = 70, Width = 70, Top = 40, DialogResult = DialogResult.OK };
            NOTconfirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(NOTconfirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = NOTconfirmation;

            return prompt.ShowDialog() == DialogResult.OK ? SomeString : "";
        }
    }
}
