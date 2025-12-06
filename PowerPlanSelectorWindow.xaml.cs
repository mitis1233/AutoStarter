using System.Collections.Generic;
using System.Windows;

namespace AutoStarter
{
    public partial class PowerPlanSelectorWindow : Window
    {
        public MainWindow.PowerPlan? SelectedPlan { get; private set; }

        public PowerPlanSelectorWindow(List<MainWindow.PowerPlan>? plans)
        {
            InitializeComponent();
            PowerPlanListBox.ItemsSource = plans;
            if (plans != null && plans.Count > 0)
            {
                PowerPlanListBox.SelectedIndex = 0;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (PowerPlanListBox.SelectedItem is MainWindow.PowerPlan selectedPlan)
            {
                SelectedPlan = selectedPlan;
                DialogResult = true;
            }
            else
            {
                DialogResult = false;
            }
            Close();
        }
    }
}
