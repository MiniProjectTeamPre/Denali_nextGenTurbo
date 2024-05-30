using Emgu.CV;
using Emgu.CV.Structure;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Denali_nextGenTurbo {
    public class ImageClass {
        public Form form { get; set; }
        public Graphics graphics { get; set; }
        public Pen pen { get; set; }
        public Point startPoint { get; set; }
        public Point stopPoint { get; set; }
        public bool penFlag { get; set; }
        public Bitmap bitmapSup { get; set; }
        public Define define { get; set; }
        public Position position { get; set; }
        public bool flagDoWork { get; set; }
        public Emgu.CV.CvEnum.TemplateMatchingType typeContain { get; set; }
        /// <summary>Value = "Open"</summary>
        public string open { get; set; }
        /// <summary>Value = "Close"</summary>
        public string close { get; set; }
        /// <summary>เก็บค่าของ scrolBar ใช้ในการเลื่อนตำแหน่งจาก position มา get image status</summary>
        public int scrollBar { get; set; }
        /// <summary>เอาไว้กำหนดให้เช็คหรือไม่เช็ค Image</summary>
        public bool flagCheck { get; set; }

        public ImageClass() {
            penFlag = false;
            define = new Define();
            position = new Position();
            typeContain = Emgu.CV.CvEnum.TemplateMatchingType.CcoeffNormed;
            open = "Open";
            close = "Close";
            flagCheck = true;
        }
        public void GetStatus(Form1 form) {
            try {
                File.ReadAllText(define.txtSaveImage);
                File.Delete(define.txtSaveImage);
                return;
            } catch { }
            Bitmap bitmap = new Bitmap(define.pathImageRefer + define.lastNamePNG);
            form.pb_imageRefer.Image = new Bitmap(bitmap);
        }
        public void SaveStatus(Form1 form, Image image_) {
            try {
                form.pb_imageRefer.Image = image_;
                Image<Bgr, byte> png = new Image<Bgr, byte>((Bitmap)form.pb_imageRefer.Image);
                png.Save(define.pathImageRefer + define.lastNamePNG);
            } catch {
                MessageBox.Show(define.errSaveImage);
                File.WriteAllText(define.txtSaveImage, string.Empty);
                form.Close();
            }
        }

        public class Position {
            public string head { get; set; }
            public string head1 { get; set; }
            public string head2 { get; set; }
            public string head3 { get; set; }
            public string head4 { get; set; }
            public string head5 { get; set; }
            public string execute { get; set; }
            public int timerKeyCtrl { get; set; }

            public Position() {
                head1 = "1";
                head2 = "2";
                head3 = "3";
                head4 = "4";
                head5 = "5";
                execute = "Execute";
            }
        }
        public class Define {
            /// <summary>Value = "../../config/ImageRefer"</summary>
            public string pathImageRefer { get; set; }
            /// <summary>Value = ".png"</summary>
            public string lastNamePNG { get; set; }
            /// <summary>Value = "_Copy.png"</summary>
            public string lastNameCopyPNG { get; set; }
            /// <summary>Value = "0.9"</summary>
            public double ContainsValue { get; set; }
            /// <summary>Value = "Please save the picture again."</summary>
            public string errSaveImage { get; set; }
            /// <summary>Value = "SaveImageError.txt"</summary>
            public string txtSaveImage { get; set; }

            public Define() {
                pathImageRefer = "../../config/ImageRefer";
                lastNamePNG = ".png";
                lastNameCopyPNG = "_Copy.png";
                ContainsValue = 0.75;
                errSaveImage = "Please save the picture again.";
                txtSaveImage = "SaveImageError.txt";
            }
        }
    }
}
