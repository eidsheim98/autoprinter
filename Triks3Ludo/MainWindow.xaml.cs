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
using WindowsInput;
using WindowsInput.Native;

namespace Triks3Ludo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        InputSimulator inputSimulator;
        bool fire = false;
        int speed;
        int printSpeed;
        string toString;
        string fromString;

        public MainWindow()
        {
            InitializeComponent();

        }


        public async void KeyCode()
        {
            try
            {
                //get speeds
                speed = int.Parse(speedInput.Text);
                printSpeed = int.Parse(printSpeedInput.Text);

                //clear errorBox
                errorBox.Clear();

                //initiate
                outPutBox.Text = "Awaiting";
                await Task.Delay(5000);

                while (fire)
                {

                    //jump one down
                    inputSimulator = new InputSimulator();
                    inputSimulator.Keyboard.KeyDown(VirtualKeyCode.DOWN);
                    outPutBox.Text = "DOWN pressed";
                    await Task.Delay(speed);

                    //copy using CTRL + C
                    inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
                    outPutBox.Text = "Copied";
                    await Task.Delay(speed);

                    //get clipboard text
                    string clipText = Clipboard.GetText();

                    //if cell is not null
                    if (clipText != "\r\n")
                    {

                        //splits clipboard and shows
                        string splitOn = splitWordInput.Text;
                        var parts = clipText.Split(new string[] { splitOn }, StringSplitOptions.None);
                        fromString = parts[0];
                        toString = splitOn + parts[1];
                        showFrom.Text = fromString;
                        showTo.Text = toString;

                        //switches to printer using ALT + TAB
                        inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.TAB);
                        await Task.Delay(speed);
                        inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                        outPutBox.Text = "Switched window focus to printer";
                        await Task.Delay(speed);

                        //selects all using CTRL + A
                        inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_A);
                        outPutBox.Text = "Selected all";
                        await Task.Delay(speed);

                        //extra safety measure
                        if (fire)
                        {
                            //paste using CTRL + V
                            Clipboard.SetText(fromString);
                            inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                            outPutBox.Text = "Pasted";
                            await Task.Delay(speed);
                            inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                            outPutBox.Text = "Enter pressed";
                            Clipboard.SetText(toString);
                            await Task.Delay(speed);
                            inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                            outPutBox.Text = "Pasted";
                            await Task.Delay(speed);

                            //print using CTRL + P
                            inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_P);
                            await Task.Delay(speed);
                            inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                            outPutBox.Text = "Printing";
                            await Task.Delay(printSpeed);

                        }

                        //switch back to excel using ALT + TAB
                        inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.TAB);
                        await Task.Delay(speed);
                        inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                        outPutBox.Text = "Switched window focus to spreadsheet";
                        await Task.Delay(speed);

                    }
                    //if cell is null
                    else
                    {
                        fire = false;
                        outPutBox.Text = "Clipboard is null - stopping";
                        await Task.Delay(5000);
                    }
                }
                //if stopbutton is pressed
                outPutBox.Text = "Stopped";
                // if an error happens
            }
            catch (Exception e)
            {
                errorBox.Text = e.ToString();
            }
        }

        public void StartOnClick(object sender, RoutedEventArgs e)
        {
            fire = true;
            KeyCode();

        }

        public void StopOnClick(object sender, RoutedEventArgs e)
        {
            fire = false;
        }


    }
}
