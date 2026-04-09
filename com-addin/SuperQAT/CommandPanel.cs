using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SuperQAT
{
    /// <summary>
    /// Floating panel with a searchable list of all QAT commands.
    /// Click any command to execute it instantly via ExecuteMso().
    /// </summary>
    public class CommandPanel : Form
    {
        private readonly ThisAddIn _addIn;
        private TextBox _searchBox;
        private ListBox _commandList;
        private Label _countLabel;
        private List<(string Name, string MsoId)> _allCommands;
        private List<(string Name, string MsoId)> _filtered;

        public CommandPanel(ThisAddIn addIn)
        {
            _addIn = addIn;
            _allCommands = MsoCommands.GetAll();
            _filtered = new List<(string, string)>(_allCommands);

            InitializeUI();
            PopulateList();
        }

        private void InitializeUI()
        {
            Text = "SuperQAT";
            Width = 340;
            Height = 600;
            MinimumSize = new Size(280, 400);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            TopMost = true;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9f);

            var headerLabel = new Label
            {
                Text = "SuperQAT",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Dock = DockStyle.Top,
                Height = 32,
                Padding = new Padding(4, 4, 0, 0),
            };

            _searchBox = new TextBox
            {
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10f),
                Height = 28,
                PlaceholderText = "Search commands...",
            };
            _searchBox.TextChanged += OnSearchChanged;

            _countLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8f),
                Padding = new Padding(4, 2, 0, 0),
            };

            _commandList = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f),
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
            };
            _commandList.Click += OnCommandClicked;
            _commandList.KeyDown += OnCommandKeyDown;

            // Add controls in reverse order (Fill last)
            Controls.Add(_commandList);
            Controls.Add(_countLabel);
            Controls.Add(_searchBox);
            Controls.Add(headerLabel);
        }

        private void PopulateList()
        {
            _commandList.BeginUpdate();
            _commandList.Items.Clear();
            foreach (var cmd in _filtered)
            {
                _commandList.Items.Add(cmd.Name);
            }
            _commandList.EndUpdate();
            _countLabel.Text = $"  {_filtered.Count} commands";
        }

        private void OnSearchChanged(object sender, EventArgs e)
        {
            var query = _searchBox.Text.Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(query))
            {
                _filtered = new List<(string, string)>(_allCommands);
            }
            else
            {
                var terms = query.Split(' ');
                _filtered = _allCommands
                    .Where(c => terms.All(t => c.Name.ToLowerInvariant().Contains(t)))
                    .ToList();
            }
            PopulateList();
        }

        private void OnCommandClicked(object sender, EventArgs e)
        {
            ExecuteSelected();
        }

        private void OnCommandKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ExecuteSelected();
                e.Handled = true;
            }
        }

        private void ExecuteSelected()
        {
            var idx = _commandList.SelectedIndex;
            if (idx < 0 || idx >= _filtered.Count) return;

            var cmd = _filtered[idx];
            _addIn.ExecuteCommand(cmd.MsoId);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Hide instead of close so it can be re-shown
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
            base.OnFormClosing(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Ctrl+F focuses the search box
            if (keyData == (Keys.Control | Keys.F))
            {
                _searchBox.Focus();
                _searchBox.SelectAll();
                return true;
            }
            // Escape closes the panel
            if (keyData == Keys.Escape)
            {
                Hide();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
