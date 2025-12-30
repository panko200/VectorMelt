using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Controls;
using YukkuriMovieMaker.Exo;
using YukkuriMovieMaker.Player.Video;
using YukkuriMovieMaker.Plugin.Effects;

namespace VectorMelt
{
    [VideoEffect("VectorMelt", ["加工"], ["data", "vector", "melt"])]
    internal class VectorMeltEffect : VideoEffectBase
    {
        public override string Label => "VectorMelt";

        // --- 基本設定 ---
        [Display(GroupName = "基本", Name = "適応", Description = "その瞬間の画像を「Iフレーム」として固定し、以降は動きベクトルだけでピクセルを動かします。")]
        [AnimationSlider("F0", "", 0, 1)]
        public Animation Freeze { get; } = new Animation(1, 0, 1);

        [Display(GroupName = "基本", Name = "更新間隔", Description = "指定したフレーム数ごとに映像を正常に戻します。0で無効。")]
        [AnimationSlider("F0", "fr", 0, 300)]
        public Animation RefreshInterval { get; } = new Animation(0, 0, 300);

        [Display(GroupName = "基本", Name = "判定間隔", Description = "何フレームに一度動きを計算するか設定します。値を増やすとコマ落ち表現になります。")]
        [AnimationSlider("F0", "fr", 1, 30)]
        public Animation UpdateInterval { get; } = new Animation(1, 1, 30);

        [Display(GroupName = "基本", Name = "ブロックサイズ", Description = "ブロックノイズっぽさを出すために、フロー計算や適用の解像度を下げるパラメータ。")]
        [AnimationSlider("F1", "px", 1, 100)]
        public Animation BlockSize { get; } = new Animation(10f, 1, 100);

        // --- 流れ制御 (XY分離 & ドリフト) ---
        [Display(GroupName = "流れ制御", Name = "強度 X", Description = "横方向の動きの適用強度")]
        [AnimationSlider("F1", "%", 0, 200)]
        public Animation IntensityX { get; } = new Animation(100, 0, 200);

        [Display(GroupName = "流れ制御", Name = "強度 Y", Description = "縦方向の動きの適用強度")]
        [AnimationSlider("F1", "%", 0, 200)]
        public Animation IntensityY { get; } = new Animation(100, 0, 200);

        [Display(GroupName = "流れ制御", Name = "ドリフト X", Description = "常に横方向に流し続ける量")]
        [AnimationSlider("F1", "px", -20, 20)]
        public Animation DriftX { get; } = new Animation(0, -20, 20);

        [Display(GroupName = "流れ制御", Name = "ドリフト Y", Description = "常に縦方向に流し続ける量")]
        [AnimationSlider("F1", "px", -20, 20)]
        public Animation DriftY { get; } = new Animation(0, -20, 20);

        // --- 応用設定 ---
        [Display(GroupName = "応用", Name = "色収差", Description = "RGBチャンネルごとに動きの強度をずらします。")]
        [AnimationSlider("F1", "%", 0, 100)]
        public Animation ColorShift { get; } = new Animation(0, 0, 100);

        [Display(GroupName = "応用", Name = "治癒", Description = "崩れた映像を徐々に元の映像に戻します。")]
        [AnimationSlider("F1", "%", 0, 100)]
        public Animation Decay { get; } = new Animation(0, 0, 100);

        [Display(GroupName = "応用", Name = "モッシュ対象", Description = "映像のどの成分を破壊するか選びます。")]
        [EnumComboBox]
        public MoshTarget Target { get; set; } = MoshTarget.All;

        [Display(GroupName = "応用", Name = "境界ぼかし", Description = "静止部分と動く部分の境界を滑らかにします。")]
        [AnimationSlider("F0", "px", 0, 50)]
        public Animation MaskSoftness { get; } = new Animation(0, 0, 50);

        [Display(GroupName = "応用", Name = "輪郭合成", Description = "現在の映像の輪郭線を薄く合成し、視認性を上げます。")]
        [AnimationSlider("F1", "%", 0, 100)]
        public Animation EdgeFactor { get; } = new Animation(0, 0, 100);

        // --- 判定設定 ---
        [Display(GroupName = "判定", Name = "処理モード", Description = "動きベクトルの計算精度")]
        [EnumComboBox]
        public MoshMode CalcMoshMode
        {
            get => calcMoshMode;
            set => Set(ref calcMoshMode, value);
        }
        private MoshMode calcMoshMode = MoshMode.Glitch;

        [Display(GroupName = "判定", Name = "移動モード", Description = "ピクセル変形時の補間方法。\nそれぞれの名前どうやって日本語にすればいいかわからんかった。とりあえず適当にOpenCvSharpの名前にした。")]
        [EnumComboBox]
        public CalcMosh CalcMoshMethodMode
        {
            get => calcMoshMethodMode;
            set => Set(ref calcMoshMethodMode, value);
        }
        private CalcMosh calcMoshMethodMode = CalcMosh.Nearest;

        [Display(GroupName = "判定", Name = "上書き", Description = "変化していない箇所を現在のフレームで置き換えるか。")]
        [EnumComboBox]
        public MaskMode CalcMaskMode
        {
            get => calcMaskMode;
            set => Set(ref calcMaskMode, value);
        }
        private MaskMode calcMaskMode = MaskMode.None;

        [Display(GroupName = "判定", Name = "変化判定", Description = "この値以下の動きは「静止」とみなして透過させます。")]
        [AnimationSlider("F1", "px", 0, 10)]
        public Animation MotionThreshold { get; } = new Animation(2.0f, 0, 100);

        // --- デバッグ ---
        [Display(GroupName = "デバッグ", Name = "動き判定を表示", Description = "静止している(白)か、動いている(黒)かの判定を可視化します。")]
        [ToggleSlider]
        public bool ShowMotionDebug { get => showMotionDebug; set => Set(ref showMotionDebug, value); }
        private bool showMotionDebug = false;

        public enum MoshMode
        {
            [Display(Name = "カクカク")]
            Glitch,
            [Display(Name = "スムース")]
            Smooth
        }

        public enum CalcMosh
        {
            [Display(Name = "Linear")]
            Linear,
            [Display(Name = "Cubic")]
            Cubic,
            [Display(Name = "Nearest")]
            Nearest
        }

        public enum MoshTarget
        {
            [Display(Name = "全体")]
            All,
            [Display(Name = "色のみ")]
            ColorOnly,
            [Display(Name = "輝度のみ")]
            LumaOnly
        }

        public enum MaskMode
        {
            [Display(Name = "なし")]
            None,
            [Display(Name = "無変化上書き")]
            ChangeMode,
            [Display(Name = "変化上書き")]
            unChangeMode
        }

        public override IEnumerable<string> CreateExoVideoFilters(int keyFrameIndex, ExoOutputDescription exoOutputDescription)
        {
            return [];
        }

        public override IVideoEffectProcessor CreateVideoEffect(IGraphicsDevicesAndContext devices)
        {
            return new VectorMeltEffectProcessor(devices, this);
        }

        protected override IEnumerable<IAnimatable> GetAnimatables() =>
            [Freeze, BlockSize, IntensityX, IntensityY, DriftX, DriftY, MotionThreshold, RefreshInterval, ColorShift, Decay, UpdateInterval, MaskSoftness, EdgeFactor];
    }
}