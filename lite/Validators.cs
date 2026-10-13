using System;
using System.Collections.Generic;

namespace FakturaLite
{
    /// <summary>ЕИК / ЕГН / номер по ЗДДС checks (same rules as the full app). They only produce warnings.</summary>
    public static class Validators
    {
        static bool AllDigits(string s) { if (s.Length == 0) return false; foreach (char c in s) if (c < '0' || c > '9') return false; return true; }

        public static bool EgnValid(string s)
        {
            if (s.Length != 10 || !AllDigits(s)) return false;
            int[] w = { 2, 4, 8, 5, 10, 9, 7, 3, 6 }; int sum = 0;
            for (int i = 0; i < 9; i++) sum += (s[i] - '0') * w[i];
            int r = sum % 11; if (r == 10) r = 0;
            return r == s[9] - '0';
        }

        static bool Nine(int[] d)
        {
            int sum = 0; for (int i = 0; i < 8; i++) sum += d[i] * (i + 1);
            int r = sum % 11;
            if (r == 10) { sum = 0; for (int i = 0; i < 8; i++) sum += d[i] * (i + 3); r = sum % 11; if (r == 10) r = 0; }
            return r == d[8];
        }

        public static bool EikValid(string s)
        {
            if (s == null) return false; s = s.Trim();
            if (!AllDigits(s)) return false;
            int[] d = new int[s.Length]; for (int i = 0; i < s.Length; i++) d[i] = s[i] - '0';
            if (s.Length == 9) return Nine(d);
            if (s.Length == 13)
            {
                if (!Nine(d)) return false;
                int r = (d[8] * 2 + d[9] * 7 + d[10] * 3 + d[11] * 5) % 11;
                if (r == 10) { r = (d[8] * 4 + d[9] * 9 + d[10] * 5 + d[11] * 7) % 11; if (r == 10) r = 0; }
                return r == d[12];
            }
            if (s.Length == 10) return EgnValid(s);
            return false;
        }

        public static List<string> PartyWarn(string eik, string vat)
        {
            List<string> w = new List<string>();
            string e = (eik ?? "").Trim(), v = (vat ?? "").Trim().ToUpperInvariant();
            if (e.Length > 0)
            {
                if (!AllDigits(e)) w.Add("ЕИК/БУЛСТАТ трябва да съдържа само цифри");
                else if (e.Length != 9 && e.Length != 10 && e.Length != 13) w.Add("ЕИК/ЕГН трябва да е 9, 10 или 13 цифри (въведени са " + e.Length + ")");
                else if (!EikValid(e)) w.Add("контролната цифра на ЕИК/ЕГН не съвпада — вероятна грешка при въвеждане");
            }
            if (v.StartsWith("BG"))
            {
                string g = v.Substring(2);
                if (!AllDigits(g) || (g.Length != 9 && g.Length != 10)) w.Add("номерът по ЗДДС трябва да е BG и 9 или 10 цифри");
                else if (g.Length == 9 && !EikValid(g)) w.Add("контролната цифра на номера по ЗДДС не съвпада");
                else if (g.Length == 9 && AllDigits(e) && e.Length >= 9 && e.Substring(0, 9) != g) w.Add("номерът по ЗДДС не съвпада с ЕИК");
            }
            return w;
        }
    }
}
