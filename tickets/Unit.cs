namespace tickets
{
    class Unit
    {
        public Unit(string art, string name, string dim, string th, string len, string wi, string note, int qty)
        {
            this.art = art;
            this.name = name;
            this.dim = dim;
            this.len = len;
            this.th = th;
            this.wi = wi;
            this.note = note;
            this.qty = qty;
        }

        public override string ToString()
        {
            return string.Format("[Art={0}, Name={1}, Dim={2}, Th={3}, Wi={4}, Len={5}, Note={6}, Qty={7}]",
                art, name, dim, th, wi, len, note, qty);
        }

        public string GetSticker()
        {
            string text = art;
            if (name != "")
                text += "\n" + name;
            if (dim != "")
                text += "\n" + dim;
            else {
                string _dim = string.Format("{0}x{1}x{2}", th, len, wi);
                if (_dim != "") {
                    _dim = _dim.Trim('x');
                    text += "\n" + _dim;
                }
            }
            if (note != "") {
                text += "\n" + note;
            }

            return text;
        }

        public string art, name, dim, note, th, len, wi;
        public int qty;
    }
}
