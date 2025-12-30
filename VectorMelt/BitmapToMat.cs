using OpenCvSharp;
using SharpGen.Runtime;
using System;
using System.Runtime.InteropServices;
using Vortice.DCommon;
using Vortice.Direct2D1;
using Vortice.DXGI;
using Vortice.Mathematics;

#nullable enable
namespace VectorMelt
{
    public static class BitmapToMat
    {
        public static Mat BitmapToOpenCvMat(ID2D1DeviceContext dc, ID2D1Bitmap bitmap)
        {
            // CPU読み取り可能なBitmapか確認し、違えば作成してコピー
            ID2D1Bitmap1? readableBitmap = null;
            var bmp1 = bitmap.QueryInterfaceOrNull<ID2D1Bitmap1>();

            if (bmp1 != null && (bmp1.Options & BitmapOptions.CpuRead) != BitmapOptions.None)
            {
                readableBitmap = bmp1;
            }
            else
            {
                bmp1?.Dispose();
                var props = new BitmapProperties1(
                    new PixelFormat(Format.B8G8R8A8_UNorm, AlphaMode.Premultiplied),
                    96f, 96f,
                    BitmapOptions.CannotDraw | BitmapOptions.CpuRead);

                readableBitmap = dc.CreateBitmap(bitmap.PixelSize, props);
                readableBitmap.CopyFromBitmap(bitmap);
            }

            // メモリマップしてOpenCVのMatにコピー
            var map = readableBitmap.Map(MapOptions.Read);
            try
            {
                // Mat作成 (BGRA)
                using var tempMat = new Mat(
                    readableBitmap.PixelSize.Height,
                    readableBitmap.PixelSize.Width,
                    MatType.CV_8UC4,
                    map.Bits,
                    map.Pitch // longキャストなどが不要な場合もありますが、エラーが出るなら (long)map.Pitch
                );
                return tempMat.Clone(); // Cloneして返す（Map外でも使えるように）
            }
            finally
            {
                readableBitmap.Unmap();
                // 一時的に作ったBitmapなら破棄
                if (readableBitmap != bitmap)
                {
                    readableBitmap.Dispose();
                }
            }
        }

        public static ID2D1Bitmap1 CreateD2DBitmapFromMat(ID2D1DeviceContext dc, Mat mat)
        {
            if (mat.Type() != MatType.CV_8UC4)
                throw new ArgumentException("Mat type must be CV_8UC4 (BGRA32)");

            // Matのデータをbyte配列にコピー（またはポインタから直接作成）
            // 安全のため一度マネージド配列を経由するか、CopyFromMemoryを使う
            var w = mat.Cols;
            var h = mat.Rows;

            var props = new BitmapProperties1(
                new PixelFormat(Format.B8G8R8A8_UNorm, AlphaMode.Premultiplied),
                96f, 96f,
                BitmapOptions.None); // TargetやCpuReadは不要、描画に使えればよい

            var bmp = dc.CreateBitmap(new SizeI(w, h), props);

            // 行ごとのステップ数(stride)
            int pitch = (int)mat.Step();

            // メモリからコピー
            bmp.CopyFromMemory(mat.Data, pitch);

            return bmp;
        }
    }
}