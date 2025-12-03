using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Transporter
{
    public partial class PlusMinusBox : UserControl
    {
        #region Свойства
        public static readonly DependencyProperty GlobalWidthProperty = DependencyProperty.Register("GlobalWidth", typeof(double), typeof(PlusMinusBox), new PropertyMetadata(null));
        public static readonly DependencyProperty TextProperty = DependencyProperty.Register("Text", typeof(string), typeof(PlusMinusBox), new PropertyMetadata(null));
        public static readonly DependencyProperty AlignProperty = DependencyProperty.Register("Align", typeof(string), typeof(PlusMinusBox), new PropertyMetadata(null));
        public static readonly DependencyProperty FillColorProperty = DependencyProperty.Register("FillColor", typeof(Brush), typeof(PlusMinusBox), new PropertyMetadata(null));
        public static readonly DependencyProperty StartColorProperty = DependencyProperty.Register("StartColor", typeof(Brush), typeof(PlusMinusBox), new PropertyMetadata(null));
        public static readonly DependencyProperty HoverColorProperty = DependencyProperty.Register("HoverColor", typeof(Brush), typeof(PlusMinusBox), new PropertyMetadata(null));

        public double GlobalWidth
        {
            get { return (double)GetValue(GlobalWidthProperty); }
            set { SetValue(GlobalWidthProperty, value); }
        }

        public string Text
        {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }

        public string Align
        {
            get { return (string)GetValue(AlignProperty); }
            set { SetValue(AlignProperty, value); }
        }

        public Brush FillColor
        {
            get { return (Brush)GetValue(FillColorProperty); }
            set { SetValue(FillColorProperty, value); }
        }

        public Brush StartColor
        {
            get { return (Brush)GetValue(StartColorProperty); }
            set { SetValue(StartColorProperty, value); }
        }

        public Brush HoverColor
        {
            get { return (Brush)GetValue(HoverColorProperty); }
            set
            {
                SetValue(HoverColorProperty, value);
                //PlusMinus_MouseEnter(this, null);
            }
        }
        #endregion



        public PlusMinusBox()
        {
            InitializeComponent();
            DataContext = this;

            if (Align == "Left") TB.HorizontalContentAlignment = HorizontalAlignment.Left;
            else if (Align == "Center") TB.HorizontalContentAlignment = HorizontalAlignment.Center;

            Text = "1";
        }



        #region Кнопки
        private void Plus_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if ((Text is null) || (Text.Trim() == ""))
                    Text = "0";

                double number = Math.Round(Convert.ToDouble(Text), 1);
                number += 1;

                Text = number.ToString();
                this.TB.SelectionStart = Text.Length;
            }
            catch
            { }
        }
        private void Minus_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if ((Text is null) || (Text.Trim() == ""))
                    Text = "0";

                double number = Math.Round(Convert.ToDouble(Text), 1);
                number -= 1;

                if (number < 0)
                    number = 0;

                Text = number.ToString();
                this.TB.SelectionStart = Text.Length;
            }
            catch
            { }
        }

        private void PlusMinus_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Animation.CreateColorAnimation((System.Windows.Shapes.Rectangle)sender, FillColor, ((SolidColorBrush)HoverColor).Color, 0.35);
        }

        private void PlusMinus_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Animation.CreateColorAnimation((System.Windows.Shapes.Rectangle)sender, FillColor, ((SolidColorBrush)StartColor).Color, 0.35);
        }

        private void Clear_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Clear();
        }
        #endregion



        #region Прочие        
        public void Clear()
        {
            Text = "";
        }

        public bool CheckIsNotEmpty()
        {
            if (!(Text is null) && (Text.Trim() != ""))
                return true;
            else return false;
        }

        public void RedVisual()
        {
            if (Text is null || Text.Trim() == "")
                Animation.CreateOpacityAnimation(IsEmpty, 7);
            else IsEmpty.Opacity = 0;
        }

        private string CheckLimits(object sender)
        {
            string output = ((TextBox)sender).Text;
            int number;

            try
            {
                number = Convert.ToInt32(output);

                if (number < 0)
                    number = 0;

                return number.ToString();
            }
            catch { }

            return output;
        }

        private void TB_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                Text = CheckLimits(sender).ToString();
            }
            catch
            { }
        }

        private void TB_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            if (!e.Text.Any(x => Char.IsDigit(x)))
                e.Handled = true;
        }
        #endregion
    }
}
