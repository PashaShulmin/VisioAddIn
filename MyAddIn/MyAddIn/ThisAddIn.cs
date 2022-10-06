using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace MyAddIn
{
    public partial class ThisAddIn
    {
        private Dictionary<string, Visio.Shape> verticesShapes;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            verticesShapes = new Dictionary<string, Visio.Shape>();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon();
            ribbon.ButtonClicked += RibbonButtonClicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { ribbon });
        }

        private void RibbonButtonClicked()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "gv files (*.gv)|*.gv|dot files (*.dot)|*.dot";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string filePath = ofd.FileName;
                    try
                    {
                        Graph graph = new Graph(filePath);
                        DropGraph(graph);
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "Ошибка, вот стек трейс", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
        }

        private void DropGraph(Graph graph)
        {
            Application.ActiveDocument.Pages.Add();

            double startX = 0, startY = 0, endX = 0.8, endY = 0.4;
            foreach (Vertex vertex in graph.GetVertices)
            {
                Visio.Shape vertexShape = Application.ActivePage.DrawOval(startX, startY, endX, endY);
                vertexShape.Text = vertex.Label;

                startX += 0.3;
                startY += 0.45;
                endX += 0.3;
                endY += 0.45;

                verticesShapes[vertex.Id] = vertexShape;
            }

            foreach (Edge edge in graph.GetEdges)
            {
                Visio.Document stencil = Application.Documents.OpenEx("Basic Flowchart Shapes (US units).vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
                Visio.Master connectorMaster = stencil.Masters.get_ItemU("Dynamic Connector");

                Visio.Shape connector = Application.ActivePage.Drop(connectorMaster, 0, 0);

                Visio.Cell srcCell = connector.get_CellsSRC(
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowXForm1D,
                    (short)Visio.VisCellIndices.vis1DBeginX);

                Visio.Shape srcShape = verticesShapes[edge.From];
                srcCell.GlueTo(srcShape.get_CellsSRC(
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowXFormOut,
                    (short)Visio.VisCellIndices.visXFormPinX));

                Visio.Cell targCell = connector.get_CellsSRC(
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowXForm1D,
                    (short)Visio.VisCellIndices.vis1DEndX);

                Visio.Shape targShape = verticesShapes[edge.To];
                targCell.GlueTo(targShape.get_CellsSRC(
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowXFormOut,
                    (short)Visio.VisCellIndices.visXFormPinX));

                if (edge.Label != "" && edge.Label != null)
                {
                    connector.Text = edge.Label;
                }

                if (edge.IsDirected)
                {
                    connector.get_Cells("EndArrow").FormulaU = "=1";
                }
            }
            try 
            {
                Application.ActivePage.Name = $"graph: {graph.Name}";
            }
            catch
            {
                MessageBox.Show(
                    "Страница с таким именем уже существует т.к. граф с таким именем уже был импортирован, " +
                    "поэтому имя страницы было заменено на стандартное.",
                    "Info", MessageBoxButtons.OK, MessageBoxIcon.Information
                    );
            }

            LayoutPage();
        }

        private void LayoutPage()
        {
            Visio.Page page = Application.ActivePage;
            page.PageSheet.get_Cells("PlaceStyle").FormulaU = "=6";
            page.PageSheet.get_Cells("RouteStyle").FormulaU = "=16";
            page.Layout();
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
