using System;

namespace HomeWork_2_Algo
{
    public class Program
    {

        static string[,] cells = new string[15, 10];
        static string[,] types = new string[15, 10];
        static int rows = 8;
        static int cols = 5;

        static void Main()
        {
            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    cells[i, j] = "";
                    types[i, j] = "unassigned";
                }
            }

            while (true)
            {
                ShowExcel();
                ShowMenu();
                Console.Write("Choose: ");
                string choice = Console.ReadLine();

                if (choice == "0")
                {
                    SaveFile();
                    break;
                }

                DoOperation(choice);
                Console.WriteLine("\nPress Enter...");
                Console.ReadLine();
            }
        }
        // نمایش صفحه اکسل 
        static void ShowExcel()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n*** EXCEL SHEET ***\n");
            Console.ResetColor();

            Console.Write("   |");
            for (int j = 0; j < cols; j++)
            {
                Console.Write("   " + (char)('A' + j) + "   |");
            }
            Console.WriteLine();

            for (int i = 0; i < rows; i++)
            {
                if (i + 1 < 10) Console.Write(" ");
                Console.Write((i + 1) + " |");

                for (int j = 0; j < cols; j++)
                {
                    string val = cells[i, j];
                    if (val.Length > 5) val = val.Substring(0, 5) + "_";

                    Console.Write(val);
                    for (int k = val.Length; k < 7; k++) Console.Write(" ");
                    Console.Write("|");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }
        //طراحی منو به صورت دلخواه برای زیبایی بیشتر برنامه 

        static void ShowMenu()
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("=== MENU ===");
            Console.ResetColor();
            Console.WriteLine("1. AssignValue  2. ClearCell    3. ClearAll");
            Console.WriteLine("4. AddRow       5. AddColumn    6. Copy");
            Console.WriteLine("7. CopyColumn   8. CopyRow      9. X");
            Console.WriteLine("10. XColumn     11. XRow        12. Multiply");
            Console.WriteLine("13. Add         14. Divide      15. Subtract");
            Console.WriteLine("16. Encrypt     17. View Cell   0. Exit");
        }
        //شروع 
        static void DoOperation(string c)
        {
            if (c == "1") AssignValue();
            else if (c == "2") ClearCell();
            else if (c == "3") ClearAll();
            else if (c == "4") AddRow();
            else if (c == "5") AddColumn();
            else if (c == "6") Copy();
            else if (c == "7") CopyColumn();
            else if (c == "8") CopyRow();
            else if (c == "9") CutCell();
            else if (c == "10") CutColumn();
            else if (c == "11") CutRow();
            else if (c == "12") Multiply();
            else if (c == "13") Add();
            else if (c == "14") Divide();
            else if (c == "15") Subtract();
            else if (c == "16") Encrypt();
            else if (c == "17") ViewCell();
            else Console.WriteLine("Wrong!");
        }
        //اضافه کردن مقدار به صفحه اکسل 
        static void AssignValue()
        {
            Console.Write("Cell: ");
            string pos = Console.ReadLine().ToUpper();
            Console.Write("Type (integer/string): ");
            string type = Console.ReadLine().ToLower();
            Console.Write("Value: ");
            string value = Console.ReadLine();

            int r = GetRow(pos);
            int c = GetCol(pos);
            if (r == -1 || c == -1) { Console.WriteLine("Wrong cell!"); return; }

            cells[r, c] = value;
            types[r, c] = type;
            Console.WriteLine("Operation is done!");
        }
        // پاکسازی کامل حانه هایه ی اکسل 
        static void ClearCell()
        {
            Console.Write("Cell: ");
            string pos = Console.ReadLine().ToUpper();
            int r = GetRow(pos);
            int c = GetCol(pos);
            if (r == -1 || c == -1) { Console.WriteLine("Wrong cell!"); return; }
            cells[r, c] = "";
            types[r, c] = "unassigned";
            Console.WriteLine("Operation is done!");
        }
        // پاک کردن همهچی
        static void ClearAll()
        {
            for (int i = 0; i < 15; i++)
                for (int j = 0; j < 10; j++)
                {
                    cells[i, j] = "";
                    types[i, j] = "unassigned";
                }
            Console.WriteLine("Operation is done!");
        }
        // اضافه کردن سطر  
        static void AddRow()
        {
            Console.Write("Row number: ");
            int num = int.Parse(Console.ReadLine());
            Console.Write("Direction (up/down): ");
            string dir = Console.ReadLine().ToLower();

            if (num < 1 || num > rows) { Console.WriteLine("Illegal position assignment!"); return; }
            if (rows >= 15) { Console.WriteLine("Out of bounds exception!"); return; }

            int at = num - 1;
            if (dir == "down") at = num;

            for (int i = rows; i > at; i--)
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j] = cells[i - 1, j];
                    types[i, j] = types[i - 1, j];
                }

            for (int j = 0; j < cols; j++)
            {
                cells[at, j] = "";
                types[at, j] = "unassigned";
            }
            rows++;
            Console.WriteLine("Operation is done!");
        }
        // اضافه کردن ستون
        static void AddColumn()
        {
            Console.Write("Column: ");
            string let = Console.ReadLine().ToUpper();
            Console.Write("Direction (left/right): ");
            string dir = Console.ReadLine().ToLower();

            int num = let[0] - 'A';
            if (num < 0 || num >= cols) { Console.WriteLine("Illegal position assignment!"); return; }
            if (cols >= 10) { Console.WriteLine("Out of bounds exception!"); return; }

            int at = num;
            if (dir == "right") at = num + 1;

            for (int j = cols; j > at; j--)
                for (int i = 0; i < rows; i++)
                {
                    cells[i, j] = cells[i, j - 1];
                    types[i, j] = types[i, j - 1];
                }

            for (int i = 0; i < rows; i++)
            {
                cells[i, at] = "";
                types[i, at] = "unassigned";
            }
            cols++;
            Console.WriteLine("Operation is done!");
        }
        // کپی 
        static void Copy()
        {
            Console.Write("From: ");
            string from = Console.ReadLine().ToUpper();
            Console.Write("To: ");
            string to = Console.ReadLine().ToUpper();

            int r1 = GetRow(from), c1 = GetCol(from);
            int r2 = GetRow(to), c2 = GetCol(to);
            if (r1 == -1 || c1 == -1 || r2 == -1 || c2 == -1) { Console.WriteLine("Wrong cell!"); return; }

            cells[r2, c2] = cells[r1, c1];
            types[r2, c2] = types[r1, c1];
            Console.WriteLine("Operation is done!");
        }
        // کپی کردن ستون
        static void CopyColumn()
        {
            Console.Write("From column: ");
            int c1 = Console.ReadLine().ToUpper()[0] - 'A';
            Console.Write("To column: ");
            int c2 = Console.ReadLine().ToUpper()[0] - 'A';

            if (c1 < 0 || c1 >= cols || c2 < 0 || c2 >= cols) { Console.WriteLine("Illegal position assignment!"); return; }

            for (int i = 0; i < rows; i++)
            {
                cells[i, c2] = cells[i, c1];
                types[i, c2] = types[i, c1];
            }
            Console.WriteLine("Operation is done!");
        }
        // کپی کردن سطر
        static void CopyRow()
        {
            Console.Write("From row: ");
            int r1 = int.Parse(Console.ReadLine()) - 1;
            Console.Write("To row: ");
            int r2 = int.Parse(Console.ReadLine()) - 1;

            if (r1 < 0 || r1 >= rows || r2 < 0 || r2 >= rows) { Console.WriteLine("Illegal position assignment!"); return; }

            for (int j = 0; j < cols; j++)
            {
                cells[r2, j] = cells[r1, j];
                types[r2, j] = types[r1, j];
            }
            Console.WriteLine("Operation is done!");
        }
        // برش خانه
        static void CutCell()
        {
            Console.Write("From: ");
            string from = Console.ReadLine().ToUpper();
            Console.Write("To: ");
            string to = Console.ReadLine().ToUpper();

            int r1 = GetRow(from), c1 = GetCol(from);
            int r2 = GetRow(to), c2 = GetCol(to);
            if (r1 == -1 || c1 == -1 || r2 == -1 || c2 == -1) { Console.WriteLine("Wrong cell!"); return; }

            cells[r2, c2] = cells[r1, c1];
            types[r2, c2] = types[r1, c1];
            cells[r1, c1] = "";
            types[r1, c1] = "unassigned";
            Console.WriteLine("Operation is done!");
        }
        // برش ستون 
        static void CutColumn()
        {
            Console.Write("From column: ");
            int c1 = Console.ReadLine().ToUpper()[0] - 'A';
            Console.Write("To column: ");
            int c2 = Console.ReadLine().ToUpper()[0] - 'A';

            if (c1 < 0 || c1 >= cols || c2 < 0 || c2 >= cols) { Console.WriteLine("Illegal position assignment!"); return; }

            for (int i = 0; i < rows; i++)
            {
                cells[i, c2] = cells[i, c1];
                types[i, c2] = types[i, c1];
                cells[i, c1] = "";
                types[i, c1] = "unassigned";
            }
            Console.WriteLine("Operation is done!");
        }
        //برش سطر 
        static void CutRow()
        {
            Console.Write("From row: ");
            int r1 = int.Parse(Console.ReadLine()) - 1;
            Console.Write("To row: ");
            int r2 = int.Parse(Console.ReadLine()) - 1;

            if (r1 < 0 || r1 >= rows || r2 < 0 || r2 >= rows) { Console.WriteLine("Illegal position assignment!"); return; }

            for (int j = 0; j < cols; j++)
            {
                cells[r2, j] = cells[r1, j];
                types[r2, j] = types[r1, j];
                cells[r1, j] = "";
                types[r1, j] = "unassigned";
            }
            Console.WriteLine("Operation is done!");
        }
        //برای انجام عملیات ضرب
        static void Multiply()
        {
            Console.Write("First cell: ");
            string c1 = Console.ReadLine().ToUpper();
            Console.Write("Second cell: ");
            string c2 = Console.ReadLine().ToUpper();
            Console.Write("Output cell: ");
            string out1 = Console.ReadLine().ToUpper();

            int r1 = GetRow(c1), col1 = GetCol(c1);
            int r2 = GetRow(c2), col2 = GetCol(c2);
            int ro = GetRow(out1), colo = GetCol(out1);

            if (r1 == -1 || col1 == -1 || r2 == -1 || col2 == -1 || ro == -1 || colo == -1) { Console.WriteLine("Wrong cell!"); return; }

            string t1 = types[r1, col1], t2 = types[r2, col2];
            if (t1 == "unassigned" || t2 == "unassigned") { Console.WriteLine("Illegal operation! Unassigned cell!"); return; }

            if (t1 == "string" && t2 == "string") { Console.WriteLine("Illegal operation! String - String operation is not allowed!"); return; }

            if (t1 == "integer" && t2 == "integer")
            {
                cells[ro, colo] = (int.Parse(cells[r1, col1]) * int.Parse(cells[r2, col2])).ToString();
                types[ro, colo] = "integer";
            }
            else
            {
                string str = (t1 == "string") ? cells[r1, col1] : cells[r2, col2];
                int num = (t1 == "integer") ? int.Parse(cells[r1, col1]) : int.Parse(cells[r2, col2]);

                string res = "";
                if (num > 0)
                {
                    for (int i = 0; i < num; i++) res += str;
                }
                else if (num < 0)
                {
                    string rev = "";
                    for (int i = str.Length - 1; i >= 0; i--) rev += str[i];
                    for (int i = 0; i < Math.Abs(num); i++) res += rev;
                }
                cells[ro, colo] = res;
                types[ro, colo] = "string";
            }
            Console.WriteLine("Operation is done!");
        }
        //برای انجام عملیات جمع
        static void Add()
        {
            Console.Write("First cell: ");
            string c1 = Console.ReadLine().ToUpper();
            Console.Write("Second cell: ");
            string c2 = Console.ReadLine().ToUpper();
            Console.Write("Third cell (Enter to skip): ");
            string c3 = Console.ReadLine().ToUpper();
            Console.Write("Output cell: ");
            string out1 = Console.ReadLine().ToUpper();

            int r1 = GetRow(c1), col1 = GetCol(c1);
            int r2 = GetRow(c2), col2 = GetCol(c2);
            int ro = GetRow(out1), colo = GetCol(out1);

            if (r1 == -1 || col1 == -1 || r2 == -1 || col2 == -1 || ro == -1 || colo == -1)
            {
                Console.WriteLine("Wrong cell!");
                return;
            }

            bool has3 = c3 != "";
            int r3 = 0, col3 = 0;
            if (has3)
            {
                r3 = GetRow(c3);
                col3 = GetCol(c3);
                if (r3 == -1 || col3 == -1)
                {
                    Console.WriteLine("Wrong cell!");
                    return;
                }
            }

            string t1 = types[r1, col1], t2 = types[r2, col2], t3 = has3 ? types[r3, col3] : "integer";
            if (t1 == "unassigned" || t2 == "unassigned" || (has3 && t3 == "unassigned"))
            {
                Console.WriteLine("Illegal operation! Unassigned cell!");
                return;
            }

            bool str = (t1 == "string" || t2 == "string" || (has3 && t3 == "string"));

            if (str)
            {
                Console.WriteLine("Concatenation operation will be applied!");
                Console.Write("Please select letter case (up/low): ");
                string mode = Console.ReadLine().ToLower();

                string res = cells[r1, col1] + cells[r2, col2];
                if (has3) res += cells[r3, col3];

                cells[ro, colo] = (mode == "low") ? res.ToLower() : res.ToUpper();
                types[ro, colo] = "string";
            }
            else
            {
                int sum = int.Parse(cells[r1, col1]) + int.Parse(cells[r2, col2]);
                if (has3) sum += int.Parse(cells[r3, col3]);
                cells[ro, colo] = sum.ToString();
                types[ro, colo] = "integer";
            }
            Console.WriteLine("Operation is done!");
        }
        //برای انجام عملیات تقسیم 
        static void Divide()
        {
            Console.Write("First cell: ");
            string c1 = Console.ReadLine().ToUpper();
            Console.Write("Second cell: ");
            string c2 = Console.ReadLine().ToUpper();
            Console.Write("Output cell: ");
            string out1 = Console.ReadLine().ToUpper();

            int r1 = GetRow(c1), col1 = GetCol(c1);
            int r2 = GetRow(c2), col2 = GetCol(c2);
            int ro = GetRow(out1), colo = GetCol(out1);

            if (r1 == -1 || col1 == -1 || r2 == -1 || col2 == -1 || ro == -1 || colo == -1) { Console.WriteLine("Wrong cell!"); return; }

            string t1 = types[r1, col1], t2 = types[r2, col2];
            if (t1 == "unassigned" || t2 == "unassigned") { Console.WriteLine("Illegal operation! Unassigned cell!"); return; }

            if (t1 == "string" && t2 == "string") { Console.WriteLine("Illegal operation! String - String operation is not allowed!"); return; }

            if (t1 == "integer" && t2 == "integer")
            {
                int divisor = int.Parse(cells[r2, col2]);
                if (divisor == 0)
                {
                    Console.WriteLine("Illegal operation! Division by zero!");
                    return;
                }
                cells[ro, colo] = (int.Parse(cells[r1, col1]) / divisor).ToString();
                types[ro, colo] = "integer";
            }
            else
            {
                string str = (t1 == "string") ? cells[r1, col1] : cells[r2, col2];
                int num = (t1 == "integer") ? int.Parse(cells[r1, col1]) : int.Parse(cells[r2, col2]);

                if (num == 0)
                {
                    Console.WriteLine("Illegal operation! Division by zero!");
                    return;
                }

                int k = str.Length / Math.Abs(num);
                string res = "";
                if (num > 0)
                    for (int i = 0; i < k; i++) res += str[i];
                else
                    for (int i = str.Length - k; i < str.Length; i++) res += str[i];

                cells[ro, colo] = res;
                types[ro, colo] = "string";
            }
            Console.WriteLine("Operation is done!");
        }
        // برای انجام عملیات تفریق
        static void Subtract()
        {
            Console.Write("First cell: ");
            string c1 = Console.ReadLine().ToUpper();
            Console.Write("Second cell: ");
            string c2 = Console.ReadLine().ToUpper();
            Console.Write("Output cell: ");
            string out1 = Console.ReadLine().ToUpper();

            int r1 = GetRow(c1), col1 = GetCol(c1);
            int r2 = GetRow(c2), col2 = GetCol(c2);
            int ro = GetRow(out1), colo = GetCol(out1);

            if (r1 == -1 || col1 == -1 || r2 == -1 || col2 == -1 || ro == -1 || colo == -1) { Console.WriteLine("Wrong cell!"); return; }

            string t1 = types[r1, col1], t2 = types[r2, col2];
            if (t1 == "unassigned" || t2 == "unassigned") { Console.WriteLine("Illegal operation! Unassigned cell!"); return; }

            if (t1 == "integer" && t2 == "integer")
            {
                cells[ro, colo] = (int.Parse(cells[r1, col1]) - int.Parse(cells[r2, col2])).ToString();
                types[ro, colo] = "integer";
            }
            else if (t1 == "string" && t2 == "string")
            {
                cells[ro, colo] = cells[r1, col1].Replace(cells[r2, col2], "");
                types[ro, colo] = "string";
            }
            else
            {
                string str = (t1 == "string") ? cells[r1, col1] : cells[r2, col2];
                int asc = (t1 == "integer") ? int.Parse(cells[r1, col1]) : int.Parse(cells[r2, col2]);

                if (asc < 33 || asc > 126) { Console.WriteLine("Illegal operation! The integer parameter should be in [33, 126]"); return; }

                string res = "";
                for (int i = 0; i < str.Length; i++)
                    if (str[i] != (char)asc) res += str[i];

                cells[ro, colo] = res;
                types[ro, colo] = "string";
            }
            Console.WriteLine("Operation is done!");
        }
        //رمز گذاری 
        static void Encrypt()
        {
            Console.Write("First cell: ");
            string c1 = Console.ReadLine().ToUpper();
            Console.Write("Second cell: ");
            string c2 = Console.ReadLine().ToUpper();
            Console.Write("Output cell: ");
            string out1 = Console.ReadLine().ToUpper();

            int r1 = GetRow(c1), col1 = GetCol(c1);
            int r2 = GetRow(c2), col2 = GetCol(c2);
            int ro = GetRow(out1), colo = GetCol(out1);

            if (r1 == -1 || col1 == -1 || r2 == -1 || col2 == -1 || ro == -1 || colo == -1) { Console.WriteLine("Wrong cell!"); return; }

            string t1 = types[r1, col1], t2 = types[r2, col2];

            if (t1 == "unassigned" || t2 == "unassigned")
            {
                Console.WriteLine("Illegal operation! Unassigned cell!");
                return;
            }

            if ((t1 == "string" && t2 == "string") || (t1 == "integer" && t2 == "integer"))
            {
                Console.WriteLine("Illegal operation! Both cells must have different types!");
                return;
            }

            string str = (t1 == "string") ? cells[r1, col1] : cells[r2, col2];
            int shift = (t1 == "integer") ? int.Parse(cells[r1, col1]) : int.Parse(cells[r2, col2]);

            if (shift < -20 || shift > 30) { Console.WriteLine("Illegal operation! The integer parameter should be in [-20, 30]"); return; }

            string res = "";
            for (int i = 0; i < str.Length; i++) res += (char)(str[i] + shift);

            cells[ro, colo] = res;
            types[ro, colo] = "string";
            Console.WriteLine("Operation is done!");
            Console.WriteLine("Encrypted value: " + res);
        }
        // مشاهده ی خانه های اکسل
        static void ViewCell()
        {
            Console.Write("Cell: ");
            string pos = Console.ReadLine().ToUpper();
            int r = GetRow(pos), c = GetCol(pos);
            if (r == -1 || c == -1) { Console.WriteLine("Wrong cell!"); return; }
            Console.WriteLine(cells[r, c]);
        }
        //سطر
        static int GetRow(string p)
        {
            if (p == null || p.Length < 2) return -1;
            p = p.ToUpper();
            string rp = "";
            for (int i = 1; i < p.Length; i++) rp += p[i];
            int r;
            if (!int.TryParse(rp, out r)) return -1;
            r--;
            if (r < 0 || r >= rows) return -1;
            return r;
        }
        //ستون
        static int GetCol(string p)
        {
            if (p == null || p.Length < 1) return -1;
            int c = char.ToUpper(p[0]) - 'A';
            if (c < 0 || c >= cols) return -1;
            return c;
        }
        //ذخیره کردن فایل 
        static void SaveFile()
        {
            System.IO.StreamWriter f = new System.IO.StreamWriter("spreadsheet.txt");
            f.Write("   |");
            for (int j = 0; j < cols; j++) f.Write("   " + (char)('A' + j) + "   |");
            f.WriteLine();

            for (int i = 0; i < rows; i++)
            {
                if (i + 1 < 10) f.Write(" ");
                f.Write((i + 1) + " |");
                for (int j = 0; j < cols; j++)
                {
                    string v = cells[i, j];
                    if (v.Length > 5) v = v.Substring(0, 5) + "_";
                    f.Write(v);
                    for (int k = v.Length; k < 7; k++) f.Write(" ");
                    f.Write("|");
                }
                f.WriteLine();
            }
            f.Close();
            Console.WriteLine("Saved!");
        }

    }
}
