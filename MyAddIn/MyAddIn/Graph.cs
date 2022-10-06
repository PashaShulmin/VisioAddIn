using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyAddIn
{
    internal class Graph
    {
        private HashSet<Vertex> vertices;
        private List<Edge> edges;

        public List<Edge> GetEdges { get { return edges; } }
        public HashSet<Vertex> GetVertices { get { return vertices; } }

        public string Name { get; set; }

        public bool IsDirected { get; set; }

        public Graph(string path)
        {
            vertices = new HashSet<Vertex>();
            edges = new List<Edge>();

            using (StreamReader sr = new StreamReader(path))
            {
                string[] firstLine = sr.ReadLine().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (firstLine[0] == "graph")
                {
                    IsDirected = false;
                }
                else if (firstLine[0] == "digraph")
                {
                    IsDirected = true;
                }
                else
                {
                    throw new ArgumentException("Неизвестный вид графа. Граф должен быть ненаправленным" +
                        " (вы должны написать просто graph) или направленным (вы должны написать digraph)");
                }

                if (firstLine[1] != "{")
                {
                    for (int i = 1; firstLine[i] != "{"; i++)
                    {
                        Name += firstLine[i];
                        if (i != firstLine.Length - 1)
                        {
                            Name += " ";
                        }
                    }
                    Name = Name.Trim();
                }
                else
                {
                    Name = "";
                }

                FillGraph(sr);
            }
        }

        private void FillGraph(StreamReader sr)
        {
            string line;

            while ((line = sr.ReadLine()) != "}")
            {
                string[] lineSplit = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                switch (lineSplit.Length)
                {
                    case 1:
                        vertices.Add(new Vertex(lineSplit[0]));
                        break;
                    case 2:
                        vertices.Add(new Vertex(lineSplit[0], GetLabel(lineSplit)));
                        break;
                    case 3:
                        CheckDirectionAndArgs(lineSplit);
                        edges.Add(new Edge(IsDirected, lineSplit[0], lineSplit[2]));
                        break;
                    case 4:
                        CheckDirectionAndArgs(lineSplit);
                        edges.Add(new Edge(GetLabel(lineSplit), IsDirected, lineSplit[0], lineSplit[2]));
                        break;
                    default:
                        break;
                }
            }
        }

        private void CheckDirectionAndArgs(string[] lineSplit)
        {
            if (IsDirected && lineSplit[1][lineSplit[1].Length - 1] == '-')
            {
                throw new ArgumentException("В ориентированном графе все рёбра должны быть направленными");
            }
            if (!IsDirected && lineSplit[1][lineSplit[1].Length - 1] == '>')
            {
                throw new ArgumentException("В неориентированном графе не должно быть направленных рёбер");
            }

            IEnumerable<string> ids = vertices.Select(v => v.Id);
            if (!ids.Contains(lineSplit[0]) || !ids.Contains(lineSplit[2]))
            {
                throw new ArgumentException("При описании ребра следует использовать только названия вершин, описанных ранее");
            }
        }

        public string GetLabel(string[] lineSplit)
        {
            int len = lineSplit.Length;
            int leftIndex = 0, rightIndex = 0;

            for (int i = 0; i < lineSplit[len - 1].Length; i++)
            {
                if (lineSplit[len - 1][i] == '"')
                {
                    leftIndex = i;
                    break;
                }
            }

            for (int i = lineSplit[len - 1].Length - 1; i >= 0; i--)
            {
                if (lineSplit[len - 1][i] == '"')
                {
                    rightIndex = i;
                    break;
                }
            }

            string label = lineSplit[len - 1].Substring(leftIndex + 1, rightIndex - leftIndex - 1);
            return label;
        }
    }
}
