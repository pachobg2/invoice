using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace FakturaLite
{
    /// <summary>JSON number kept as its original text, so a round trip never changes a value.</summary>
    public sealed class JNum
    {
        public string Raw;
        public JNum(string raw) { Raw = raw; }
        public double D { get { double d; return double.TryParse(Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0; } }
    }

    /// <summary>JSON object that remembers key order (so the file keeps its layout).</summary>
    public sealed class JObj
    {
        public List<string> Keys = new List<string>();
        public Dictionary<string, object> Map = new Dictionary<string, object>();
        public object this[string key]
        {
            get { object v; return Map.TryGetValue(key, out v) ? v : null; }
            set { if (!Map.ContainsKey(key)) Keys.Add(key); Map[key] = value; }
        }
        public bool Has(string key) { return Map.ContainsKey(key); }
        public string Str(string key)
        {
            object v = this[key];
            if (v == null) return "";
            JNum n = v as JNum; if (n != null) return n.Raw;
            return v as string ?? (v is bool ? ((bool)v ? "true" : "false") : "");
        }
        public double Num(string key)
        {
            object v = this[key]; JNum n = v as JNum; if (n != null) return n.D;
            string s = v as string; double d;
            if (s != null && double.TryParse(s.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            return 0;
        }
        public List<object> Arr(string key)
        {
            List<object> a = this[key] as List<object>;
            if (a == null) { a = new List<object>(); this[key] = a; }
            return a;
        }
        public JObj Obj(string key)
        {
            JObj o = this[key] as JObj;
            if (o == null) { o = new JObj(); this[key] = o; }
            return o;
        }
    }

    public static class Json
    {
        public static JNum Num(double d) { return new JNum(d.ToString("0.##########", CultureInfo.InvariantCulture)); }
        public static JNum Num(long n) { return new JNum(n.ToString(CultureInfo.InvariantCulture)); }

        /* ------------------------------ parser ------------------------------ */
        public static object Parse(string text)
        {
            int i = 0;
            object v = ReadValue(text, ref i);
            SkipWs(text, ref i);
            if (i != text.Length) throw new FormatException("Неочакван текст в края на JSON файла.");
            return v;
        }

        static void SkipWs(string s, ref int i) { while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n' || s[i] == '﻿')) i++; }

        static object ReadValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) throw new FormatException("Непълен JSON файл.");
            char c = s[i];
            if (c == '{')
            {
                JObj o = new JObj(); i++; SkipWs(s, ref i);
                if (s[i] == '}') { i++; return o; }
                while (true)
                {
                    SkipWs(s, ref i);
                    string k = ReadString(s, ref i);
                    SkipWs(s, ref i);
                    if (s[i] != ':') throw new FormatException("Очаква се ':' в JSON.");
                    i++;
                    o[k] = ReadValue(s, ref i);
                    SkipWs(s, ref i);
                    if (s[i] == ',') { i++; continue; }
                    if (s[i] == '}') { i++; return o; }
                    throw new FormatException("Очаква се ',' или '}' в JSON.");
                }
            }
            if (c == '[')
            {
                List<object> a = new List<object>(); i++; SkipWs(s, ref i);
                if (s[i] == ']') { i++; return a; }
                while (true)
                {
                    a.Add(ReadValue(s, ref i));
                    SkipWs(s, ref i);
                    if (s[i] == ',') { i++; continue; }
                    if (s[i] == ']') { i++; return a; }
                    throw new FormatException("Очаква се ',' или ']' в JSON.");
                }
            }
            if (c == '"') return ReadString(s, ref i);
            if (string.CompareOrdinal(s, i, "true", 0, 4) == 0) { i += 4; return true; }
            if (string.CompareOrdinal(s, i, "false", 0, 5) == 0) { i += 5; return false; }
            if (string.CompareOrdinal(s, i, "null", 0, 4) == 0) { i += 4; return null; }
            int st = i;
            while (i < s.Length && "+-0123456789.eE".IndexOf(s[i]) >= 0) i++;
            if (i == st) throw new FormatException("Неочакван знак в JSON: " + c);
            return new JNum(s.Substring(st, i - st));
        }

        static string ReadString(string s, ref int i)
        {
            if (s[i] != '"') throw new FormatException("Очаква се текст в JSON.");
            i++;
            StringBuilder sb = new StringBuilder();
            while (true)
            {
                char c = s[i++];
                if (c == '"') return sb.ToString();
                if (c != '\\') { sb.Append(c); continue; }
                c = s[i++];
                switch (c)
                {
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'u': sb.Append((char)int.Parse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture)); i += 4; break;
                    default: sb.Append(c); break; // " \ /
                }
            }
        }

        /* ------------------------------ writer ------------------------------ */
        public static string Write(object v)
        {
            StringBuilder sb = new StringBuilder();
            Write(sb, v);
            return sb.ToString();
        }

        static void Write(StringBuilder sb, object v)
        {
            if (v == null) { sb.Append("null"); return; }
            string s = v as string; if (s != null) { WriteString(sb, s); return; }
            JNum n = v as JNum; if (n != null) { sb.Append(n.Raw); return; }
            if (v is bool) { sb.Append((bool)v ? "true" : "false"); return; }
            JObj o = v as JObj;
            if (o != null)
            {
                sb.Append('{');
                for (int k = 0; k < o.Keys.Count; k++)
                {
                    if (k > 0) sb.Append(',');
                    WriteString(sb, o.Keys[k]); sb.Append(':'); Write(sb, o.Map[o.Keys[k]]);
                }
                sb.Append('}'); return;
            }
            List<object> a = v as List<object>;
            if (a != null)
            {
                sb.Append('[');
                for (int k = 0; k < a.Count; k++) { if (k > 0) sb.Append(','); Write(sb, a[k]); }
                sb.Append(']'); return;
            }
            WriteString(sb, v.ToString());
        }

        static void WriteString(StringBuilder sb, string s)
        {
            sb.Append('"');
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }
    }
}
