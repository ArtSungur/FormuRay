using System;

namespace FormuRay.Core
{
    public sealed class ExplorerHotkey
    {
        public bool Control { get; private set; }
        public bool Alt { get; private set; }
        public bool Shift { get; private set; }
        public char Letter { get; private set; }
        public string KeyName { get; private set; }
        public int VirtualKey { get; private set; }
        public string Canonical { get { return (Control ? "^" : "") + (Alt ? "%" : "") + (Shift ? "+" : "") + KeyName.ToLowerInvariant(); } }
        public string DisplayName { get { return (Control ? "Ctrl + " : "") + (Shift ? "Shift + " : "") + (Alt ? "Alt + " : "") + KeyName; } }
        public string ModifierLabel { get { return (Control ? "Ctrl+" : "") + (Alt ? "Alt+" : "") + (Shift ? "Shift+" : ""); } }
        public static ExplorerHotkey Parse(string text, bool extended = false)
        {
            text = (text ?? "").Trim();
            var key = new ExplorerHotkey();
            int i = 0;
            for (; i < text.Length && "^%+".IndexOf(text[i]) >= 0; i++)
            {
                switch (text[i])
                {
                    case '^': if (key.Control) throw new ArgumentException("Duplicate Ctrl modifier."); key.Control = true; break;
                    case '%': if (key.Alt) throw new ArgumentException("Duplicate Alt modifier."); key.Alt = true; break;
                    case '+': if (key.Shift) throw new ArgumentException("Duplicate Shift modifier."); key.Shift = true; break;
                }
            }
            string letter = text.Substring(i).Trim().ToUpperInvariant();
            int function;
            bool character = letter.Length == 1 && ((letter[0] >= 'A' && letter[0] <= 'Z')
                || (extended && letter[0] >= '0' && letter[0] <= '9'));
            bool fkey = extended && letter.StartsWith("F", StringComparison.Ordinal)
                && int.TryParse(letter.Substring(1), out function) && function >= 1 && function <= 24;
            if (i == 0 || (!character && !fkey))
                throw new ArgumentException(extended ? "Choose at least one modifier and A-Z, 0-9 or F1-F24." : "Choose at least one modifier and a letter A-Z.");
            key.KeyName = letter;
            key.Letter = character ? letter[0] : '\0';
            key.VirtualKey = character ? letter[0] : 111 + int.Parse(letter.Substring(1));
            if (extended && (key.VirtualKey == 123 || (key.Alt && !key.Control && !key.Shift && key.VirtualKey == 115)))
                throw new ArgumentException("F12 and Alt + F4 are reserved. Choose another shortcut.");
            return key;
        }
        public bool Matches(int letter, bool control, bool alt, bool shift)
        {
            return VirtualKey == letter && Control == control && Alt == alt && Shift == shift;
        }
    }

    public sealed class ExplorerHotkeySettings
    {
        public ExplorerHotkey Precedents { get; private set; }
        public ExplorerHotkey Dependents { get; private set; }
        public ExplorerHotkey PasteFormulas { get; private set; }
        public HighlightSettings Highlights { get; private set; }
        public ExplorerHotkeySettings(string precedents, string dependents, string pasteFormulas = "%+v", HighlightSettings highlights = null)
        {
            Highlights = highlights ?? HighlightSettings.Defaults;
            Precedents = ExplorerHotkey.Parse(precedents);
            Dependents = ExplorerHotkey.Parse(dependents);
            PasteFormulas = ExplorerHotkey.Parse(pasteFormulas, true);
            if (Precedents.Canonical == Dependents.Canonical || Precedents.Canonical == PasteFormulas.Canonical
                || Dependents.Canonical == PasteFormulas.Canonical)
                throw new ArgumentException("FormuRay shortcuts cannot be identical.");
            foreach (var shortcut in new[] { Precedents, Dependents, PasteFormulas })
                if (shortcut.Canonical == "^c" || shortcut.Canonical == "^x" || shortcut.Canonical == "^v")
                    throw new ArgumentException("Ctrl + C, Ctrl + X and Ctrl + V are reserved for standard Excel copy, cut and paste.");
        }
        public static ExplorerHotkeySettings Defaults { get { return new ExplorerHotkeySettings("+%w", "+%s"); } }
    }
}
