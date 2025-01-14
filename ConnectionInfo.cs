/* TABLE OF CONTENTS
 * 
 * FIELDS...
 * 
 * CONSTRUCTORS...
 * 
 * METHODS
 * print()
 * 
 */

using System;
using System.Windows.Forms;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class ConnectionInfo
    {
        //======================================================================================================================
        // FIELDS
        //======================================================================================================================
        public string Environ;
        public string Server;
        public string Database;
        public string Schema;
        public string Account;
        public string Username;
        public string Password;

        //======================================================================================================================
        // CONSTRUCTORS
        //======================================================================================================================
        public ConnectionInfo() 
        {
            //get it straight from the form
            ComboBox EnvironComboBox = Application.OpenForms["Form1"].Controls["EnvironComboBox"] as ComboBox;
            ComboBox ServerComboBox = Application.OpenForms["Form1"].Controls["ServerComboBox"] as ComboBox;
            ComboBox DatabaseComboBox = Application.OpenForms["Form1"].Controls["DatabaseComboBox"] as ComboBox;
            ComboBox SchemaComboBox = Application.OpenForms["Form1"].Controls["SchemaComboBox"] as ComboBox;
            ComboBox AccountComboBox = Application.OpenForms["Form1"].Controls["AccountComboBox"] as ComboBox;
            ComboBox UsernameComboBox = Application.OpenForms["Form1"].Controls["UsernameComboBox"] as ComboBox;
            ComboBox PasswordComboBox = Application.OpenForms["Form1"].Controls["PasswordComboBox"] as ComboBox;

            Environ = EnvironComboBox.Text;
            Server = ServerComboBox.Text;
            Database = DatabaseComboBox.Text;
            Schema = SchemaComboBox.Text;
            Account = AccountComboBox.Text;
            Username = UsernameComboBox.Text;
            Password = PasswordComboBox.Text;
        }

        public ConnectionInfo(string Environ_ = "", string Server_ = "", string Database_ = "", string Schema_ = "", string Account_ = "", string Username_ = "", string Password_ = "")
        {
            //let the user fill out the object
            Environ = Environ_;
            Server = Server_;
            Database = Database_;
            Schema = Schema_;
            Account = Account_;
            Username = Username_;
            Password = Password_;
        }

        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public void print()
        {
            //print the contents of the object
            if (Environ != null && Environ != "") { Console.WriteLine("Environ: " + Environ); }
            if (Server != null && Server != "") { Console.WriteLine("Server: " + Server); }
            if (Database != null && Database != "") { Console.WriteLine("Database: " + Database); }
            if (Schema != null && Schema != "") { Console.WriteLine("Schema: " + Schema); }
            if (Account != null && Account != "") { Console.WriteLine("Account: " + Account); }
            if (Username != null && Username != "") { Console.WriteLine("Username: " + Username); }
            //if (Password != null && Password != "") { Console.WriteLine("Password: " + Password); }
            Console.WriteLine("");
        }
    }
}
