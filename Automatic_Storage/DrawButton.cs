using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Automatic_Storage
{
    class DrawButton
    {
        public enum ControlState { Hover, Normal, Pressed }
        public class RoundButton : Button
        {

            private int radius;//半径 
            public int Radius
            {
                set
                {
                    radius = value;
                    this.Invalidate();
                }
                get
                {
                    return radius;
                }
            }
            public ControlState ControlState { get; set; }
            protected override void OnMouseEnter(EventArgs e)//鼠标进入时
            {
                base.OnMouseEnter(e);
                ControlState = ControlState.Hover;//正常
            }
            protected override void OnMouseLeave(EventArgs e)//鼠标离开
            {
                base.OnMouseLeave(e);
                ControlState = ControlState.Normal;//正常
            }
            protected override void OnMouseDown(MouseEventArgs e)//鼠标按下
            {
                base.OnMouseDown(e);
                if (e.Button == MouseButtons.Left && e.Clicks == 1)//鼠标左键且点击次数为1
                {
                    ControlState = ControlState.Pressed;//按下的状态
                }
            }
            protected override void OnMouseUp(MouseEventArgs e)//鼠标弹起
            {
                base.OnMouseUp(e);
                if (e.Button == MouseButtons.Left && e.Clicks == 1)
                {
                    if (ClientRectangle.Contains(e.Location))//控件区域包含鼠标的位置
                    {
                        ControlState = ControlState.Hover;
                    }
                    else
                    {
                        ControlState = ControlState.Normal;
                    }
                }
            }
            public RoundButton()
            {
                Radius = 15;
                this.FlatStyle = FlatStyle.Flat;
                this.FlatAppearance.BorderSize = 0;
                this.ControlState = ControlState.Normal;
                this.SetStyle(
                 ControlStyles.UserPaint |  //控件自行绘制，而不使用操作系统的绘制
                 ControlStyles.AllPaintingInWmPaint | //忽略擦出的消息，减少闪烁。
                 ControlStyles.OptimizedDoubleBuffer |//在缓冲区上绘制，不直接绘制到屏幕上，减少闪烁。
                 ControlStyles.ResizeRedraw | //控件大小发生变化时，重绘。                  
                 ControlStyles.SupportsTransparentBackColor, true);//支持透明背景颜色
            }

            //重写OnPaint
            protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
            {
                base.OnPaint(e);
                base.OnPaintBackground(e);
                e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
                e.Graphics.CompositingQuality = CompositingQuality.HighQuality;

                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;

                Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
                var path = GetRoundedRectPath(rect, radius);

                this.Region = new Region(path);

                Color baseColor = this.BackColor;

                using (SolidBrush b = new SolidBrush(baseColor))
                {
                    e.Graphics.FillPath(b, path);
                    Font fo = new Font("微軟正黑體", 18F);//指定預設字體與大小
                    Brush brush = new SolidBrush(Color.White);
                    StringFormat gs = new StringFormat();
                    gs.Alignment = StringAlignment.Center; //居中
                    gs.LineAlignment = StringAlignment.Center;//垂直居中
                    e.Graphics.DrawString(this.Text, fo, brush, rect, gs);
                }
            }
            private GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
            {
                int diameter = radius;
                Rectangle arcRect = new Rectangle(rect.Location, new Size(diameter, diameter));
                GraphicsPath path = new GraphicsPath();
                path.AddArc(arcRect, 180, 90);
                arcRect.X = rect.Right - diameter;
                path.AddArc(arcRect, 270, 90);
                arcRect.Y = rect.Bottom - diameter;
                path.AddArc(arcRect, 0, 90);
                arcRect.X = rect.Left;
                path.AddArc(arcRect, 90, 90);
                path.CloseFigure();
                return path;
            }

            protected override void OnSizeChanged(EventArgs e)
            {
                base.OnSizeChanged(e);
            }
        }
    }
}
