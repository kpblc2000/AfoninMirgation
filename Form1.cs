using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Useful_FunctionsCsh
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.KeyPress += textBox1_KeyPress;
            textBox2.KeyPress += textBox2_KeyPress;
            textBox3.KeyPress += textBox3_KeyPress;
            textBox4.KeyPress += textBox4_KeyPress;
            textBox5.KeyPress += textBox5_KeyPress;
            textBox6.KeyPress += textBox6_KeyPress;
            textBox7.KeyPress += textBox7_KeyPress;
            textBox8.KeyPress += textBox8_KeyPress;
            textBox9.KeyPress += textBox9_KeyPress;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Formclosing(object sender, EventArgs e)
        {
           // textBox1.
           // SaveFileDialog saveFileDialog = new SaveFileDialog();
           // saveFileDialog.Title = "Сохранение изменений";

           // saveFileDialog.ShowDialog();
           //// saveFileDialog.FileOk += SaveFileDialog_FileOk;

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }
        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.SelectAll();
        }
        private void textBox3_Enter(object sender, EventArgs e)
        {
            textBox3.SelectAll();
        }
        private void textBox4_Enter(object sender, EventArgs e)
        {
            textBox4.SelectAll();
        }
        private void textBox5_Enter(object sender, EventArgs e)
        {
            textBox5.SelectAll();
        }
        private void textBox6_Enter(object sender, EventArgs e)
        {
            textBox6.SelectAll();
        }
        private void textBox7_Enter(object sender, EventArgs e)
        {
            textBox7.SelectAll();
        }
        private void textBox8_Enter(object sender, EventArgs e)
        {
            textBox8.SelectAll();
        }
        private void textBox9_Enter(object sender, EventArgs e)
        {
            textBox9.SelectAll();
        }
        private void textBox10_Enter(object sender, EventArgs e)
        {
            textBox10.SelectAll();
        }
        private void textBox11_Enter(object sender, EventArgs e)
        {
            textBox11.SelectAll();
        }
        private void textBox12_Enter(object sender, EventArgs e)
        {
            textBox12.SelectAll();
        }

       
        public bool nonNumberEntered = false;
        public void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            // Initialize the flag to false.
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            } 
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if (( e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }





        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            // Initialize the flag to false.
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode<Keys.D0 || e.KeyCode> Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode<Keys.NumPad0 || e.KeyCode> Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            } 
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
                {
                    nonNumberEntered = true;
                }

//добавляем, что можно нажимать Enter
                if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
                {
                    nonNumberEntered = false;
                }
}

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox7_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox8_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox9_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }

        private void textBox11_KeyDown(object sender, KeyEventArgs e)
        {
            nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }

            //добавляем, что можно нажимать Enter
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Decimal))
            {
                nonNumberEntered = false;
            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
           // if ((nonNumberEntered == true) || (textBox1.Text.IndexOf('.') != -1))
                if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
            else if ((e.KeyChar == ((char)Keys.Enter))|| (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox2.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox4.Focus(); }
            else if (e.KeyChar == ((char)Keys.Right))
            { textBox6.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
            else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox3.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox1.Focus(); }
            else if (e.KeyChar == ((char)Keys.Right))
            { textBox6.Focus(); }
            else
            { e.Handled = false; }

            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox4.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox2.Focus(); }
            else if (e.KeyChar == ((char)Keys.Right))
            { textBox6.Focus(); }
            else
            { e.Handled = false; }

            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox5.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox3.Focus(); }
            else if (e.KeyChar == ((char)Keys.Right))
            { textBox6.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox1.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox4.Focus(); }
            else if (e.KeyChar == ((char)Keys.Right))
            { textBox9.Focus(); }
            else
            { e.Handled = false; }

            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox7.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox9.Focus(); }
            else if (e.KeyChar == ((char)Keys.Left))
            { textBox1.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox8.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox6.Focus(); }
            else if (e.KeyChar == ((char)Keys.Left))
            { textBox5.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox9.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox7.Focus(); }
            else if (e.KeyChar == ((char)Keys.Left))
            { textBox5.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
           else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox6.Focus(); }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox8.Focus(); }
            else if (e.KeyChar == ((char)Keys.Left))
            { textBox5.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (nonNumberEntered == true)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
            else if ((e.KeyChar == ((char)Keys.Enter)) || (e.KeyChar == ((char)Keys.Down)))
            {
                SelectNextControl(sender as TextBox, true, true, true, true);
                textBox12.Focus();
            }
            else if (e.KeyChar == ((char)Keys.Up))
            { textBox10.Focus(); }
            else if (e.KeyChar == ((char)Keys.Left))
            { textBox5.Focus(); }
            else
            { e.Handled = false; }
            if (e.KeyChar == ',')
            {
                // запятую заменим точкой
                e.KeyChar = '.';
            }
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
