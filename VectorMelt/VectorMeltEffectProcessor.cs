using OpenCvSharp;
using SharpGen.Runtime;
using System;
using System.Collections.Generic;
using System.Numerics;
using Vortice.DCommon;  // 追加: PixelFormat, AlphaModeのため
using Vortice.Direct2D1;
using Vortice.Direct2D1.Effects;
using Vortice.DXGI;     // 追加: Format等のため
using Vortice.Mathematics;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Player.Video;
using static VectorMelt.VectorMeltEffect;

#nullable enable
namespace VectorMelt
{
    internal class VectorMeltEffectProcessor : IVideoEffectProcessor, IDrawable, IDisposable
    {
        private readonly DisposeCollector disposer = new DisposeCollector();
        private readonly IGraphicsDevicesAndContext devices;
        private readonly VectorMeltEffect item;
        private readonly AffineTransform2D wrap;
        public ID2D1Image Output { get; }

        private ID2D1Image? inputCurrent;

        // --- バッファ群 ---
        private Mat? _accumulatorMat;
        private Mat? _prevInputMat;

        private Mat _grayPrev = new Mat(); private Mat _grayCurr = new Mat();
        private Mat _grayPrevSmall = new Mat(); private Mat _grayCurrSmall = new Mat();
        private Mat _flowSmall = new Mat(); private Mat _flowLarge = new Mat();
        private Mat _flowInt = new Mat();

        private Mat _mapX = new Mat(); private Mat _mapY = new Mat();
        private Mat _mapX_R = new Mat(); private Mat _mapY_R = new Mat();
        private Mat _mapX_B = new Mat(); private Mat _mapY_B = new Mat();
        private Mat _gridX = new Mat(); private Mat _gridY = new Mat();

        private Mat _magnitude = new Mat();
        private Mat _mask = new Mat();
        private Mat _maskBlur = new Mat();
        private Mat _maskDebugColor = new Mat();

        private Mat _chB = new Mat(); private Mat _chG = new Mat();
        private Mat _chR = new Mat(); private Mat _chA = new Mat();
        private Mat _warpedB = new Mat(); private Mat _warpedG = new Mat();
        private Mat _warpedR = new Mat();

        private Mat _yuvCurr = new Mat(); private Mat _yuvWarped = new Mat(); private Mat _bgrResult = new Mat();
        private Mat _yC = new Mat(); private Mat _crC = new Mat(); private Mat _cbC = new Mat();
        private Mat _yW = new Mat(); private Mat _crW = new Mat(); private Mat _cbW = new Mat();

        private Mat _edges = new Mat();
        private Mat _edgesColor = new Mat();

        private ID2D1Bitmap1? _outputBitmapCache;

        public VectorMeltEffectProcessor(IGraphicsDevicesAndContext devices, VectorMeltEffect item)
        {
            this.devices = devices;
            this.item = item;
            this.wrap = new AffineTransform2D((ID2D1DeviceContext)devices.DeviceContext);
            this.disposer.Collect(this.wrap);

            // Outputの解放漏れ対策
            this.Output = this.wrap.Output;
            this.disposer.Collect(this.Output);
        }

        public DrawDescription Update(EffectDescription effectDescription)
        {
            if (this.inputCurrent == null)
            {
                this.wrap.SetInput(0, null, true);
                return effectDescription.DrawDescription;
            }

            var dc = this.devices.DeviceContext;
            int frame = effectDescription.ItemPosition.Frame;
            int duration = effectDescription.ItemDuration.Frame;
            int fps = effectDescription.FPS;

            // --- パラメータ ---
            double freezeVal = this.item.Freeze.GetValue(frame, duration, fps);
            bool isFreeze = freezeVal >= 0.5;
            int interval = (int)this.item.RefreshInterval.GetValue(frame, duration, fps);
            int updateInterval = Math.Max(1, (int)this.item.UpdateInterval.GetValue(frame, duration, fps));
            double blockSize = this.item.BlockSize.GetValue(frame, duration, fps);
            var calcMoshMode = this.item.CalcMoshMode;
            bool showDebug = this.item.ShowMotionDebug;

            bool isForceReset = (interval > 0) && (frame > 0) && (frame % interval == 0);
            bool isUpdateFrame = (frame == 0) || (frame % updateInterval == 0);

            // offsetを取得
            using var currentBitmap = ConvertToBitmap(dc, this.inputCurrent, out var offset);

            // 初期化
            if (_accumulatorMat == null || _accumulatorMat.IsDisposed ||
                currentBitmap.PixelSize.Width != _accumulatorMat.Cols ||
                currentBitmap.PixelSize.Height != _accumulatorMat.Rows)
            {
                ResetBuffers(dc, currentBitmap);

                // 初回も位置補正が必要
                this.wrap.TransformMatrix = Matrix3x2.CreateTranslation(offset);
                this.wrap.SetInput(0, currentBitmap, true);
                return effectDescription.DrawDescription;
            }

            // --- Datamosh処理 ---
            if (!isFreeze || isForceReset)
            {
                // ★ 変更: 内部メソッド呼び出し
                using var tempMat = BitmapToOpenCvMat(dc, currentBitmap);
                tempMat.CopyTo(_accumulatorMat);
                tempMat.CopyTo(_prevInputMat);
                _mask.SetTo(0);
            }
            else if (isUpdateFrame)
            {
                // ★ 変更: 内部メソッド呼び出し
                using var matCurr = BitmapToOpenCvMat(dc, currentBitmap);

                if (_prevInputMat != null)
                {
                    Cv2.CvtColor(_prevInputMat, _grayPrev, ColorConversionCodes.BGRA2GRAY);
                    Cv2.CvtColor(matCurr, _grayCurr, ColorConversionCodes.BGRA2GRAY);

                    double scale = 1.0 / Math.Max(1.0, blockSize);
                    OpenCvSharp.Size smallSize = new OpenCvSharp.Size(
                        Math.Max(1, (int)(matCurr.Cols * scale)),
                        Math.Max(1, (int)(matCurr.Rows * scale)));

                    Cv2.Resize(_grayPrev, _grayPrevSmall, smallSize, 0, 0, InterpolationFlags.Linear);
                    Cv2.Resize(_grayCurr, _grayCurrSmall, smallSize, 0, 0, InterpolationFlags.Linear);

                    Cv2.CalcOpticalFlowFarneback(
                        _grayPrevSmall, _grayCurrSmall, _flowSmall,
                        0.5, 3, 15, 3, 5, 1.2, 0);

                    if (calcMoshMode == MoshMode.Glitch)
                    {
                        _flowSmall.ConvertTo(_flowInt, MatType.CV_16SC2);
                        _flowInt.ConvertTo(_flowSmall, MatType.CV_32FC2);
                    }

                    Cv2.Resize(_flowSmall, _flowLarge, matCurr.Size(), 0, 0, InterpolationFlags.Nearest);

                    ProcessDatamosh(frame, duration, fps, matCurr);

                    matCurr.CopyTo(_prevInputMat);
                }
            }

            // --- 出力 ---
            if (_outputBitmapCache != null)
            {
                _outputBitmapCache.Dispose();
                _outputBitmapCache = null;
            }

            if (showDebug)
            {
                if (!_mask.Empty())
                {
                    var maskToShow = (this.item.MaskSoftness.GetValue(frame, duration, fps) > 0) ? _maskBlur : _mask;
                    Cv2.CvtColor(maskToShow, _maskDebugColor, ColorConversionCodes.GRAY2BGRA);
                    // ★ 変更: 内部メソッド呼び出し
                    _outputBitmapCache = CreateD2DBitmapFromMat(dc, _maskDebugColor);
                }
                else
                {
                    // ★ 変更: 内部メソッド呼び出し
                    _outputBitmapCache = CreateD2DBitmapFromMat(dc, _accumulatorMat);
                }
            }
            else
            {
                if (_accumulatorMat != null)
                {
                    // ★ 変更: 内部メソッド呼び出し
                    _outputBitmapCache = CreateD2DBitmapFromMat(dc, _accumulatorMat);
                }
            }

            if (_outputBitmapCache != null)
            {
                // ズレた座標を元に戻す
                this.wrap.TransformMatrix = Matrix3x2.CreateTranslation(offset);
                this.wrap.SetInput(0, _outputBitmapCache, true);
            }

            return effectDescription.DrawDescription;
        }

        private void ProcessDatamosh(int frame, int duration, int fps, Mat matCurr)
        {
            if (_accumulatorMat == null) return;

            // --- パラメータ取得 ---
            double intensityX = this.item.IntensityX.GetValue(frame, duration, fps) / 100.0 * 5.0;
            double intensityY = this.item.IntensityY.GetValue(frame, duration, fps) / 100.0 * 5.0;
            float driftX = (float)this.item.DriftX.GetValue(frame, duration, fps);
            float driftY = (float)this.item.DriftY.GetValue(frame, duration, fps);
            double colorShift = this.item.ColorShift.GetValue(frame, duration, fps) / 100.0;
            double edgeFactor = this.item.EdgeFactor.GetValue(frame, duration, fps) / 100.0;
            double decay = this.item.Decay.GetValue(frame, duration, fps) / 100.0;
            var calcMaskMode = this.item.CalcMaskMode;
            double motionThreshold = this.item.MotionThreshold.GetValue(frame, duration, fps);
            int maskSoftness = (int)this.item.MaskSoftness.GetValue(frame, duration, fps);
            var targetMode = this.item.Target;

            var interpolation = this.item.CalcMoshMethodMode switch
            {
                CalcMosh.Cubic => InterpolationFlags.Cubic,
                CalcMosh.Linear => InterpolationFlags.Linear,
                _ => InterpolationFlags.Nearest
            };

            // ========================================================
            // 1. マスク作成とベクトル消去
            // ========================================================
            using var flowX = _flowLarge.ExtractChannel(0);
            using var flowY = _flowLarge.ExtractChannel(1);
            Cv2.Magnitude(flowX, flowY, _magnitude);

            Cv2.Threshold(_magnitude, _mask, motionThreshold, 255, ThresholdTypes.BinaryInv);
            _mask.ConvertTo(_mask, MatType.CV_8U);
            _flowLarge.SetTo(new Scalar(0, 0), _mask);

            // ========================================================
            // 2. リマップ (変形適用)
            // ========================================================
            using var warped = new Mat(_accumulatorMat.Size(), _accumulatorMat.Type());

            if (colorShift > 0.01)
            {
                float factorR = (float)(1.0 - colorShift);
                float factorB = (float)(1.0 + colorShift);

                UpdateMaps(_flowLarge, (float)intensityX, (float)intensityY, driftX, driftY, _mapX, _mapY);
                UpdateMaps(_flowLarge, (float)(intensityX * factorR), (float)(intensityY * factorR), driftX, driftY, _mapX_R, _mapY_R);
                UpdateMaps(_flowLarge, (float)(intensityX * factorB), (float)(intensityY * factorB), driftX, driftY, _mapX_B, _mapY_B);

                Cv2.ExtractChannel(_accumulatorMat, _chB, 0);
                Cv2.ExtractChannel(_accumulatorMat, _chG, 1);
                Cv2.ExtractChannel(_accumulatorMat, _chR, 2);
                Cv2.ExtractChannel(_accumulatorMat, _chA, 3);

                Cv2.Remap(_chR, _warpedR, _mapX_R, _mapY_R, interpolation, BorderTypes.Reflect);
                Cv2.Remap(_chG, _warpedG, _mapX, _mapY, interpolation, BorderTypes.Reflect);
                Cv2.Remap(_chB, _warpedB, _mapX_B, _mapY_B, interpolation, BorderTypes.Reflect);
                using var warpedA = new Mat();
                Cv2.Remap(_chA, warpedA, _mapX, _mapY, interpolation, BorderTypes.Reflect);

                Cv2.Merge(new[] { _warpedB, _warpedG, _warpedR, warpedA }, warped);
            }
            else
            {
                UpdateMaps(_flowLarge, (float)intensityX, (float)intensityY, driftX, driftY, _mapX, _mapY);
                Cv2.Remap(_accumulatorMat, warped, _mapX, _mapY, interpolation, BorderTypes.Reflect);
            }

            // ========================================================
            // 3. マスク合成 (Blend)
            // ========================================================
            if (calcMaskMode != MaskMode.None)
            {
                Mat maskToUse;
                if (maskSoftness > 0)
                {
                    int ksize = maskSoftness * 2 + 1;
                    Cv2.GaussianBlur(_mask, _maskBlur, new OpenCvSharp.Size(ksize, ksize), 0);
                    maskToUse = _maskBlur;
                }
                else
                {
                    maskToUse = _mask;
                }

                if (calcMaskMode == MaskMode.ChangeMode)
                {
                    BlendWithMask(warped, matCurr, maskToUse);
                }
                else if (calcMaskMode == MaskMode.unChangeMode)
                {
                    using var invMask = new Mat();
                    Cv2.BitwiseNot(maskToUse, invMask);
                    BlendWithMask(warped, matCurr, invMask);
                }
            }
            else if (this.item.ShowMotionDebug && maskSoftness > 0)
            {
                int ksize = maskSoftness * 2 + 1;
                Cv2.GaussianBlur(_mask, _maskBlur, new OpenCvSharp.Size(ksize, ksize), 0);
            }

            // Target, Edge, Decay...
            if (targetMode != MoshTarget.All)
            {
                Cv2.CvtColor(matCurr, _yuvCurr, ColorConversionCodes.BGRA2BGR);
                Cv2.CvtColor(_yuvCurr, _yuvCurr, ColorConversionCodes.BGR2YCrCb);
                Cv2.CvtColor(warped, _yuvWarped, ColorConversionCodes.BGRA2BGR);
                Cv2.CvtColor(_yuvWarped, _yuvWarped, ColorConversionCodes.BGR2YCrCb);

                Cv2.ExtractChannel(_yuvCurr, _yC, 0); Cv2.ExtractChannel(_yuvCurr, _crC, 1); Cv2.ExtractChannel(_yuvCurr, _cbC, 2);
                Cv2.ExtractChannel(_yuvWarped, _yW, 0); Cv2.ExtractChannel(_yuvWarped, _crW, 1); Cv2.ExtractChannel(_yuvWarped, _cbW, 2);

                Mat yFinal, crFinal, cbFinal;
                if (targetMode == MoshTarget.ColorOnly) { yFinal = _yC; crFinal = _crW; cbFinal = _cbW; }
                else { yFinal = _yW; crFinal = _crC; cbFinal = _cbC; }

                Cv2.Merge(new[] { yFinal, crFinal, cbFinal }, _yuvWarped);
                Cv2.CvtColor(_yuvWarped, _bgrResult, ColorConversionCodes.YCrCb2BGR);

                using var b = _bgrResult.ExtractChannel(0);
                using var g = _bgrResult.ExtractChannel(1);
                using var r = _bgrResult.ExtractChannel(2);
                using var a = warped.ExtractChannel(3);
                Cv2.Merge(new[] { b, g, r, a }, warped);
            }

            if (edgeFactor > 0.0)
            {
                using var grayForEdge = new Mat();
                Cv2.CvtColor(matCurr, grayForEdge, ColorConversionCodes.BGRA2GRAY);
                Cv2.Canny(grayForEdge, _edges, 100, 200);
                Cv2.CvtColor(_edges, _edgesColor, ColorConversionCodes.GRAY2BGRA);
                Cv2.AddWeighted(warped, 1.0, _edgesColor, edgeFactor, 0, warped);
            }

            if (decay > 0.001)
            {
                Cv2.AddWeighted(warped, 1.0 - decay, matCurr, decay, 0, warped);
            }

            warped.CopyTo(_accumulatorMat);
        }

        private void BlendWithMask(Mat warped, Mat curr, Mat mask8u)
        {
            using var warpedF = new Mat(); using var currF = new Mat();
            using var maskF = new Mat(); using var maskInvF = new Mat();

            warped.ConvertTo(warpedF, MatType.CV_32F);
            curr.ConvertTo(currF, MatType.CV_32F);
            mask8u.ConvertTo(maskF, MatType.CV_32F, 1.0 / 255.0);
            using var maskF4 = new Mat();
            Cv2.CvtColor(maskF, maskF4, ColorConversionCodes.GRAY2BGRA);
            Cv2.Subtract(new Scalar(1.0, 1.0, 1.0, 1.0), maskF4, maskInvF);
            Cv2.Multiply(warpedF, maskInvF, warpedF);
            Cv2.Multiply(currF, maskF4, currF);
            Cv2.Add(warpedF, currF, warpedF);
            warpedF.ConvertTo(warped, MatType.CV_8U);
        }

        private void UpdateMaps(Mat flow, float intensityX, float intensityY, float driftX, float driftY, Mat destMapX, Mat destMapY)
        {
            int rows = flow.Rows;
            int cols = flow.Cols;
            if (_gridX.Rows != rows || _gridX.Cols != cols)
            {
                _gridX.Dispose(); _gridY.Dispose();
                _gridX = new Mat(rows, cols, MatType.CV_32FC1);
                _gridY = new Mat(rows, cols, MatType.CV_32FC1);
                using var rowX = new Mat(1, cols, MatType.CV_32FC1);
                unsafe { float* p = (float*)rowX.DataPointer; for (int x = 0; x < cols; x++) p[x] = x; }
                Cv2.Repeat(rowX, rows, 1, _gridX);
                using var colY = new Mat(rows, 1, MatType.CV_32FC1);
                unsafe { for (int y = 0; y < rows; y++) colY.Set<float>(y, 0, y); }
                Cv2.Repeat(colY, 1, cols, _gridY);
            }
            if (destMapX.Rows != rows || destMapX.Cols != cols) { destMapX.Create(rows, cols, MatType.CV_32FC1); }
            if (destMapY.Rows != rows || destMapY.Cols != cols) { destMapY.Create(rows, cols, MatType.CV_32FC1); }

            using var flowX = flow.ExtractChannel(0);
            using var flowY = flow.ExtractChannel(1);
            Cv2.Multiply(flowX, intensityX, flowX);
            Cv2.Add(flowX, driftX, flowX);
            Cv2.Subtract(_gridX, flowX, destMapX);
            Cv2.Multiply(flowY, intensityY, flowY);
            Cv2.Add(flowY, driftY, flowY);
            Cv2.Subtract(_gridY, flowY, destMapY);
        }

        private void ResetBuffers(ID2D1DeviceContext dc, ID2D1Bitmap1 source)
        {
            _accumulatorMat?.Dispose();
            _prevInputMat?.Dispose();
            // ★ 変更: 内部メソッド呼び出し
            using var temp = BitmapToOpenCvMat(dc, source);
            _accumulatorMat = temp.Clone();
            _prevInputMat = temp.Clone();
        }

        private ID2D1Bitmap1 ConvertToBitmap(ID2D1DeviceContext dc, ID2D1Image image, out Vector2 offset)
        {
            var localBounds = dc.GetImageLocalBounds(image);
            offset = new Vector2(localBounds.Left, localBounds.Top);

            int w = (int)Math.Ceiling(localBounds.Right - localBounds.Left);
            int h = (int)Math.Ceiling(localBounds.Bottom - localBounds.Top);
            if (w <= 0) w = 1; if (h <= 0) h = 1;

            float dpiX, dpiY; dc.GetDpi(out dpiX, out dpiY);
            var bmpProps = new BitmapProperties1(
                new Vortice.DCommon.PixelFormat(Vortice.DXGI.Format.B8G8R8A8_UNorm, Vortice.DCommon.AlphaMode.Premultiplied),
                dpiX, dpiY, BitmapOptions.Target);

            var bmp = dc.CreateBitmap(new SizeI(w, h), bmpProps);
            var oldTarget = dc.Target;
            dc.Target = bmp;
            dc.BeginDraw();
            dc.Clear(new Vortice.Mathematics.Color4(0, 0, 0, 0));
            dc.DrawImage(image, new Vector2(-localBounds.Left, -localBounds.Top));
            dc.EndDraw();
            dc.Target = oldTarget;
            return bmp;
        }

        public void ClearInput() => this.wrap.SetInput(0, null, true);
        public void SetInput(ID2D1Image? input) => this.inputCurrent = input;

        public void Dispose()
        {
            this.disposer.Dispose();
            _accumulatorMat?.Dispose(); _prevInputMat?.Dispose();
            _grayPrev.Dispose(); _grayCurr.Dispose();
            _grayPrevSmall.Dispose(); _grayCurrSmall.Dispose();
            _flowSmall.Dispose(); _flowLarge.Dispose(); _flowInt.Dispose();
            _mapX.Dispose(); _mapY.Dispose();
            _mapX_R.Dispose(); _mapY_R.Dispose();
            _mapX_B.Dispose(); _mapY_B.Dispose();
            _gridX.Dispose(); _gridY.Dispose();
            _magnitude.Dispose(); _mask.Dispose(); _maskBlur.Dispose(); _maskDebugColor.Dispose();
            _chB.Dispose(); _chG.Dispose(); _chR.Dispose(); _chA.Dispose();
            _warpedB.Dispose(); _warpedG.Dispose(); _warpedR.Dispose();
            _yuvCurr.Dispose(); _yuvWarped.Dispose(); _bgrResult.Dispose();
            _yC.Dispose(); _crC.Dispose(); _cbC.Dispose();
            _yW.Dispose(); _crW.Dispose(); _cbW.Dispose();
            _edges.Dispose(); _edgesColor.Dispose();
            _outputBitmapCache?.Dispose();
        }

        // --- 以下統合されたBitmapToMatのメソッド ---

        private Mat BitmapToOpenCvMat(ID2D1DeviceContext dc, ID2D1Bitmap bitmap)
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
                    map.Pitch
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

        private ID2D1Bitmap1 CreateD2DBitmapFromMat(ID2D1DeviceContext dc, Mat mat)
        {
            if (mat.Type() != MatType.CV_8UC4)
                throw new ArgumentException("Mat type must be CV_8UC4 (BGRA32)");

            var w = mat.Cols;
            var h = mat.Rows;

            var props = new BitmapProperties1(
                new PixelFormat(Format.B8G8R8A8_UNorm, AlphaMode.Premultiplied),
                96f, 96f,
                BitmapOptions.None);

            var bmp = dc.CreateBitmap(new SizeI(w, h), props);

            int pitch = (int)mat.Step();

            bmp.CopyFromMemory(mat.Data, pitch);

            return bmp;
        }
    }
}