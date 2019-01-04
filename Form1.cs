using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geometry;
using System.IO;
using System.Diagnostics;
using ESRI.ArcGIS.Output;
using ESRI.ArcGIS.Controls;
namespace ArcGISTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSrcMxdPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择数据源MXD所在文件夹";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                this.textBox1.Text = dialog.SelectedPath;
            }
        }

        private void btnTemplatePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择模板MXD所在文件夹";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                this.textBox2.Text = dialog.SelectedPath;
            }

        }

        private void btnSure_Click(object sender, EventArgs e)
        {
            string folderPath = this.textBox1.Text.Trim();
            string templateFolderPath = this.textBox2.Text.Trim();
            if (!string.IsNullOrEmpty(folderPath))
            {
                DirectoryInfo myDir = new DirectoryInfo(folderPath);
                FileInfo[] _fileList = myDir.GetFiles();
                for (int i = 0; i < _fileList.Length; i++)
                {
                    string aa = _fileList[i].Name.ToString();
                    string bb = System.IO.Path.GetExtension(aa).ToString();
                    if (System.IO.Path.GetExtension(_fileList[i].Name).Contains(".mxd") || System.IO.Path.GetExtension(_fileList[i].Name).Contains(".mxt"))
                    {
                        string mxdName = _fileList[i].Name;
                        Console.WriteLine(mxdName.Substring(3));
                        string mxdPath = folderPath + "\\" + mxdName;
                        string templatePath = templateFolderPath + "\\宗地草图" + mxdName.Substring(3);
                        string savePath = templatePath;
                        ChangelayoutAndSave(mxdPath, templatePath, savePath);
                    }
                }
                MessageBox.Show("执行成功！");
            }
            //  changelayout();
        }
        public void ChangelayoutAndSave(string mxdPath, string templatePath, string savePath)
        {
           
            //   IMapDocument pDoc = new MapDocumentClass();
            IMxdContents pMxdC;
            IMapDocument pMapDocument = new MapDocumentClass();
            pMapDocument.Open(mxdPath,"");
            pMxdC = pMapDocument.PageLayout as IMxdContents;

            IMap pMap = pMxdC.ActiveView.FocusMap; //this.axPageLayoutControl1.ActiveView.FocusMap;
            IPageLayout pPageLayout = pMxdC.PageLayout;//this.axPageLayoutControl1.PageLayout;
            //读取新模板
            IMapDocument pNewDoc = new MapDocumentClass();
            pNewDoc.Open(templatePath,"");
            IMap pTempMap;
            IPageLayout pTempPagelayout = pNewDoc.PageLayout;
            pTempMap = pNewDoc.get_Map(0);
            IPage pTempPage = pTempPagelayout.Page;

            IPage pCurPage = pPageLayout.Page;

            //替换单位
            pCurPage.Units = pTempPage.Units;
            //exchange page orientation
            pCurPage.Orientation = pTempPage.Orientation;
            //替换页面尺寸
            Double dWidth = 0;
            Double dHeight = 0;
            pTempPage.QuerySize(out dWidth, out dHeight);
            pCurPage.PutCustomSize(dWidth, dHeight);

            //删除当前Layout中除了mapframe外的所有elements
            IGraphicsContainer pGraphicsCont;
            IElement pElement;
            pGraphicsCont = pPageLayout as IGraphicsContainer;
            pGraphicsCont.Reset();
            pElement = pGraphicsCont.Next();
            IMapFrame pMapFrame = null;
            IElement pMapFrameElement = null;
            while (pElement != null)
            {

                if (pElement is IMapFrame)
                {
                    // Console.WriteLine(pElement.Geometry.GeometryType.ToString() + '0');
                    pMapFrameElement = pElement;
                    pMapFrame = pElement as IMapFrame;
                    pMapFrame.Border = null;
                }
                else
                {
                    pGraphicsCont.DeleteElement(pElement);
                    pGraphicsCont.Reset();
                }
                pElement = pGraphicsCont.Next();
            }

            //遍历模板的PageLayout中的所有元素，并且替换当前PageLayout中的所有元素
            IGraphicsContainer pTempGraphicsCont;
            pTempGraphicsCont = pTempPagelayout as IGraphicsContainer;
            pTempGraphicsCont.Reset();
            pElement = pTempGraphicsCont.Next();
            IArray pArray;
            pArray = new ESRI.ArcGIS.esriSystem.Array();
            while (pElement != null)
            {
                if (pElement is IMapFrame)
                {
                    pMapFrameElement.Geometry = pElement.Geometry;
                }
                else
                {
                    if (pElement is IMapSurroundFrame)
                    {
                        IMapSurround pTempMapSurround;
                        IMapSurroundFrame pTempMapSurroundFrame = pElement as IMapSurroundFrame;
                        pTempMapSurroundFrame.MapFrame = pMapFrame;
                        pTempMapSurround = pTempMapSurroundFrame.MapSurround;
                        pMap.AddMapSurround(pTempMapSurround);
                    }

                    pArray.Add(pElement);
                }
                pElement = pTempGraphicsCont.Next();
            }

            int pElementCount = pArray.Count;
            //将模板PageLayout中的其它元素（除了MapFrameElement和MapSurroundFrame外的元素）添加到当前PageLayout中去
            for (int i = 0; i < pElementCount; i++)
            {
                pGraphicsCont.AddElement(pArray.get_Element(pElementCount - 1 - i) as IElement, 0);

                // this.axPageLayoutControl1.ActiveView.Refresh();
            }
            //     pMxdC.ActiveView.Refresh();
            pNewDoc.Close();
            pMapDocument = new MapDocumentClass();
            pMapDocument.New(savePath);
            // IActiveView pActiveView = this.axPageLayoutControl1.ActiveView.FocusMap as IActiveView;
            pMapDocument.ReplaceContents(pMxdC);
            pMapDocument.Save(true, true);
            pMapDocument.Close();
        }

        private void btnMxdPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择数据源MXD所在文件夹";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                this.textBox3.Text = dialog.SelectedPath;
            }
        }

        private void btnSavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择地图导出路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                this.textBox1.Text = dialog.SelectedPath;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string folderPath = this.textBox3.Text.Trim();
            string saveFolderPath = this.textBox4.Text.Trim();
            if (!string.IsNullOrEmpty(folderPath) && !string.IsNullOrEmpty(saveFolderPath))
            {
                DirectoryInfo myDir = new DirectoryInfo(folderPath);
                FileInfo[] _fileList = myDir.GetFiles();
                for (int i = 0; i < _fileList.Length; i++)
                {
                    string aa = _fileList[i].Name.ToString();
                    string bb = System.IO.Path.GetExtension(aa).ToString();
                    if (System.IO.Path.GetExtension(_fileList[i].Name).Contains(".mxd"))
                    {
                        string mxdName = _fileList[i].Name;
                     //   Console.WriteLine(mxdName.Substring(3));
                        string mxdPath = folderPath + "\\" + mxdName;
                        string saveMxdPath = saveFolderPath + "\\new_" + mxdName.Split('.')[0];
                       
                        OutputMap(mxdPath, saveMxdPath);
                    }
                }
                MessageBox.Show("执行成功！");
            }
        }
        private void OutputMap(string mxdPath,string saveMxdPath)
        {
            try
            {
                if (txtResolution.Text.Trim() == "" && cbxExpention.Text.Trim() == "")
                {
                    MessageBox.Show("保存类型和分辨率不能为空！");
                    return;
                }

             //   string fpath;
                IActiveView pActiveView;
                IExport pExport;
                IEnvelope pPixelBoundsEnv;
                int iOutputResolution;
                int iScreenResolution;
                int hDC;
                IMxdContents pMxdC;
                IMapDocument pMapDocument = new MapDocumentClass();
                pMapDocument.Open(mxdPath, "");
                pMxdC = pMapDocument.PageLayout as IMxdContents;

            //    IMap pMap = pMxdC.ActiveView.FocusMap; //this.axPageLayoutControl1.ActiveView.FocusMap;
             //   IPageLayout pPageLayout = pMxdC.PageLayout;//this.axPageLayoutControl1.PageLayout;
            //    AxPageLayoutControl pPageLayoutControl = pMapDocument as AxPageLayoutControl;
                pActiveView = pMapDocument.ActiveView;
                pExport = new ExportJPEGClass();
              //  SaveFileDialog saveFileDialog1 = new SaveFileDialog();
              //  saveFileDialog1.InitialDirectory = Application.StartupPath;
              //  saveFileDialog1.Filter =string.Format("{0} files (*.{0})|*.{0}",cbxExpention.Text.Trim());
                string fpath = saveMxdPath + "."+cbxExpention.Text.Trim();
                //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                //{
                   // fpath = saveFileDialog1.FileName.Trim();
                pExport.ExportFileName = fpath;
                iScreenResolution = 96;
                iOutputResolution = 300;
                pExport.Resolution = iOutputResolution;
                tagRECT pExportFrame;
                pExportFrame = pActiveView.ExportFrame;
                tagRECT exportRECT;
                exportRECT.left = 0;
                exportRECT.top = 0;
                exportRECT.right = pActiveView.ExportFrame.right * (iOutputResolution /
                iScreenResolution);
                exportRECT.bottom = pActiveView.ExportFrame.bottom * (iOutputResolution /
                iScreenResolution);
                pPixelBoundsEnv = new EnvelopeClass();
                pPixelBoundsEnv.PutCoords(exportRECT.left, exportRECT.top,
                exportRECT.right, exportRECT.bottom);
                pExport.PixelBounds = pPixelBoundsEnv;
                hDC = pExport.StartExporting();
                pActiveView.Output(hDC, (int)pExport.Resolution, ref exportRECT, null, null);
                pExport.FinishExporting();
                pExport.Cleanup();
               // }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}