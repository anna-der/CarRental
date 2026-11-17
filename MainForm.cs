using System;
using System.Drawing;
using System.Windows.Forms;

namespace CarRentalApp
{
    public class MainForm : Form
    {
        public MainForm()
        {
            Text = "Car Rental — Система прокату автомобілів";
            Size = new Size(800, 520);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(650, 450);
            BackColor = Color.FromArgb(245, 247, 250);

            Label title = new Label
            {
                Text = "СИСТЕМА ПРОКАТУ АВТОМОБІЛІВ",
                Dock = DockStyle.Top,
                Height = 100,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 20, FontStyle.Bold),
                ForeColor = Color.FromArgb(35, 55, 80)
            };

            Label description = new Label
            {
                Text = "Оберіть розділ для роботи з базою даних",
                Dock = DockStyle.Top,
                Height = 45,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11),
                ForeColor = Color.DimGray
            };

            FlowLayoutPanel panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(30, 15, 30, 15)
            };

            panel.Controls.Add(CreateButton("Автомобілі", () =>
                new DataForm("Cars").ShowDialog(this)));

            panel.Controls.Add(CreateButton("Клієнти", () =>
                new DataForm("Clients").ShowDialog(this)));

            panel.Controls.Add(CreateButton("Оренда автомобілів", () =>
                new DataForm("Rentals").ShowDialog(this)));

            panel.Controls.Add(CreateButton("SQL-запити", () =>
                new QueryForm().ShowDialog(this)));

            Controls.Add(panel);
            Controls.Add(description);
            Controls.Add(title);
        }

        private Button CreateButton(string text, Action action)
        {
            Button button = new Button
            {
                Text = text,
                Width = 320,
                Height = 55,
                Margin = new Padding(10),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = Color.FromArgb(38, 76, 115),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;
            button.Click += (sender, e) => action();

            return button;
        }
    }
}