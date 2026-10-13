using System;
using System.Collections.Generic;
using System.Globalization;

namespace FakturaLite
{
    /// <summary>Amount in Bulgarian words, e.g. "сто двадесет и три евро и 45 евро цента" (same rules as the full app).</summary>
    public static class Words
    {
        static readonly string[] U = { "", "едно", "две", "три", "четири", "пет", "шест", "седем", "осем", "девет", "десет", "единадесет", "дванадесет", "тринадесет", "четиринадесет", "петнадесет", "шестнадесет", "седемнадесет", "осемнадесет", "деветнадесет" };
        static readonly string[] T = { "", "", "двадесет", "тридесет", "четиридесет", "петдесет", "шестдесет", "седемдесет", "осемдесет", "деветдесет" };
        static readonly string[] H = { "", "сто", "двеста", "триста", "четиристотин", "петстотин", "шестстотин", "седемстотин", "осемстотин", "деветстотин" };

        static string Fix(string w, int n, char gender)
        {
            if (gender == 'f' && n % 10 == 1 && n != 11) return "една";
            if (gender == 'm' && n % 10 == 1 && n != 11) return "един";
            if ((gender == 'f' || gender == 'm') && n % 10 == 2 && n != 12) return gender == 'm' ? "два" : "две";
            return w;
        }

        static List<string> Group(int n, char gender)
        {
            List<string> t = new List<string>();
            int h = n / 100, rest = n % 100;
            if (h > 0) t.Add(H[h]);
            if (rest > 0)
            {
                if (rest < 20) t.Add(Fix(U[rest], rest, gender));
                else { t.Add(T[rest / 10]); if (rest % 10 > 0) t.Add(Fix(U[rest % 10], rest % 10, gender)); }
            }
            return t;
        }

        class G { public List<string> Tokens; public string Scale; public G(List<string> t, string s) { Tokens = t; Scale = s; } }

        public static string IntWords(long n)
        {
            if (n == 0) return "нула";
            long bil = n / 1000000000L, mil = (n / 1000000L) % 1000, th = (n / 1000L) % 1000, un = n % 1000;
            List<G> g = new List<G>();
            if (bil > 0) g.Add(bil == 1 ? new G(new List<string>(new string[] { "един" }), "милиард") : new G(Group((int)bil, 'm'), "милиарда"));
            if (mil > 0) g.Add(mil == 1 ? new G(new List<string>(new string[] { "един" }), "милион") : new G(Group((int)mil, 'm'), "милиона"));
            if (th > 0) g.Add(th == 1 ? new G(new List<string>(), "хиляда") : new G(Group((int)th, 'f'), "хиляди"));
            if (un > 0) g.Add(new G(Group((int)un, 'n'), ""));
            G last = g[g.Count - 1];
            List<string> o = new List<string>();
            foreach (G x in g)
            {
                List<string> t = new List<string>(x.Tokens);
                if (x == last)
                {
                    if (t.Count > 1) t.Insert(t.Count - 1, "и");
                    else if (t.Count == 1 && g.Count > 1) t.Insert(0, "и");
                    else if (t.Count == 0 && g.Count > 1) o.Add("и");
                }
                o.AddRange(t);
                if (x.Scale.Length > 0) o.Add(x.Scale);
            }
            return string.Join(" ", o.ToArray());
        }

        public static string Slov(decimal v)
        {
            v = Math.Round(v, 2, MidpointRounding.AwayFromZero);
            bool neg = v < 0; if (neg) v = -v;
            long whole = (long)Math.Floor(v);
            int cents = (int)Math.Round((v - whole) * 100m, 0, MidpointRounding.AwayFromZero);
            string s = (neg ? "минус " : "") + IntWords(whole) + " евро";
            if (cents > 0) s += " и " + cents.ToString("00", CultureInfo.InvariantCulture) + " евро цента";
            return s;
        }
    }
}
