using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyAddIn
{
    internal class Vertex
    {
        public string Id { get; }

        public string Label { get; }

        public Vertex(string id, string label)
        {
            Id = id;
            Label = label;
        }

        public Vertex(string id)
        {
            Id = Label = id;
        }
    }
}
