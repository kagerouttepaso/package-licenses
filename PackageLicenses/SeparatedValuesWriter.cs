using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PackageLicenses
{
    //http://blog.xin9le.net/entry/2014/01/29/122551
    //からもってきたファイル書き込みクラス

    /// <summary>
    /// 区切り文字ファイル書き込みクラス
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public class SeparatedValuesWriter : IDisposable
    {
        /// <summary>
        /// 書き込み設定クラス
        /// </summary>
        public class SeparatedValuesWriterSetting
        {
            #region 一般的な形式のインスタンス取得
            public static SeparatedValuesWriterSetting Csv { get { return new SeparatedValuesWriterSetting(); } }
            public static SeparatedValuesWriterSetting Tsv { get { return new SeparatedValuesWriterSetting() { FieldSeparator = "\t" }; } }
            public static SeparatedValuesWriterSetting Ssv { get { return new SeparatedValuesWriterSetting() { FieldSeparator = " " }; } }
            #endregion

            #region プロパティ
            public string FieldSeparator { get; set; }   //--- フィールドの区切り文字
            public string RecordSeparator { get; set; }  //--- レコードの区切り文字
            public string TextModifier { get; set; }     //--- テキストの修飾子
            #endregion

            #region コンストラクタ
            public SeparatedValuesWriterSetting()
            {
                //--- 既定はCSV
                this.FieldSeparator = ",";
                this.RecordSeparator = Environment.NewLine;
                this.TextModifier = "\"";
            }
            #endregion
        }
        
        #region フィールド / プロパティ
        private readonly TextWriter writer = null;
        public SeparatedValuesWriterSetting Setting { get; private set; }
        #endregion


        #region コンストラクタ
        public SeparatedValuesWriter(string path, SeparatedValuesWriterSetting setting)
            : this(new StreamWriter(path), setting) { }

        public SeparatedValuesWriter(string path, bool append, SeparatedValuesWriterSetting setting)
            : this(new StreamWriter(path, append), setting) { }

        public SeparatedValuesWriter(string path, bool append, Encoding encoding, SeparatedValuesWriterSetting setting)
            : this(new StreamWriter(path, append, encoding), setting) { }

        public SeparatedValuesWriter(Stream stream, SeparatedValuesWriterSetting setting)
            : this(new StreamWriter(stream), setting) { }

        public SeparatedValuesWriter(Stream stream, Encoding encoding, SeparatedValuesWriterSetting setting)
            : this(new StreamWriter(stream, encoding), setting) { }

        public SeparatedValuesWriter(TextWriter writer, SeparatedValuesWriterSetting setting)
        {
            this.writer = writer;
            this.Setting = setting;
        }
        #endregion


        #region IDisposable メンバー
        public void Dispose()
        {
            this.Close();
        }
        #endregion


        #region 書き込み関連メソッド
        //--- 1行分のデータを一気に同期書き込み
        public void WriteLine<T>(IEnumerable<T> fields, bool quoteAlways = false)
        {
			if (fields == null)
				throw new ArgumentNullException("fields");

			string record = FormatFields(fields, quoteAlways);
			this.writer.Write(record + this.Setting.RecordSeparator);
		}

        //--- 1行分のデータを一気に非同期書き込み
        public Task WriteLineAsync<T>(IEnumerable<T> fields, bool quoteAlways = false)
		{
			string record = FormatFields(fields, quoteAlways);
			return this.writer.WriteAsync(record + this.Setting.RecordSeparator);
		}

		private string FormatFields<T>(IEnumerable<T> fields, bool quoteAlways)
		{
			if (fields == null)
				throw new ArgumentNullException("fields");

			var formated = fields.Select(x => this.FormatField(x, quoteAlways));
			var record = string.Join(this.Setting.FieldSeparator, formated);
			return record;
		}

		public void Flush()
        {
            this.writer.Flush();
        }

        public Task FlushAsync()
        {
            return this.writer.FlushAsync();
        }

        public void Close()
        {
            this.writer.Close();
        }
        #endregion


        #region 補助メソッド
        //--- フィールド文字列を整形します
        private string FormatField<T>(T field, bool quoteAlways = false)
        {
            var text = field is string ? field as string
                     : field?.ToString();
            text = text ?? string.Empty;

            if (quoteAlways || this.NeedsQuote(text))
            {
                var modifier = this.Setting.TextModifier;
                var escape = modifier + modifier;
                var builder = new StringBuilder(text);
                builder.Replace(modifier, escape);
                builder.Insert(0, modifier);
                builder.Append(modifier);
                return builder.ToString();
            }
            return text;
        }

        //--- 指定された文字列を引用符で括る必要があるかどうかを判定
        private bool NeedsQuote(string text)
        {
            return text.Contains('\r')
                || text.Contains('\n')
                || text.Contains(this.Setting.TextModifier)
                || text.Contains(this.Setting.FieldSeparator)
                || text.StartsWith("\t")
                || text.StartsWith(" ")
                || text.EndsWith("\t")
                || text.EndsWith(" ");
        }
        #endregion
    }
}
