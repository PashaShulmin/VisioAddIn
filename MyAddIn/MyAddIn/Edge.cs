using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyAddIn
{
    internal class Edge
    {
        public string Label { get; }

        public bool IsDirected { get; }

        public string From { get; set; }

        public string To { get; set; }

        public Edge(string label, bool isDirected, string from, string to)
        {
            Label = label;
            IsDirected = isDirected;
            From = from;
            To = to;
        }

        public Edge(bool isDirected, string from, string to)
        {
            Label = "";
            IsDirected = isDirected;
            From = from;
            To = to;
        }
    }
}
