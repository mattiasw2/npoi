using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NPOI.SS.UserModel;

namespace mattias1
{
    public record XlsxFile
    {
        public string FileName;
        public ImmutableList<XlsxSheet> Sheets;
        public ImmutableList<XlsxPicture> Pictures;

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public record XlsxSheet()
    {
        public int sheet;
        public string name;
        public ImmutableList<XlsxRow> Rows;

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public record XlsxRow()
    {
        public int sheet;
        public int row;
        public ImmutableList<XlsxCell> Cells;

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public record XlsxCell()
    {
        public int sheet;
        public int row;
        public int col;

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public record XlsxPicture
    {
        public string FileName;

        /**
         * Suggests a file extension for this image.
         *
         * @return the file extension.
         */
        public string SuggestFileExtension;

        /**
         * Returns the mime type for the image
         */
        public string MimeType;

        /**
         * @return the POI internal image type, 0 if unknown image type
         *
         * @see Workbook#PICTURE_TYPE_DIB
         * @see Workbook#PICTURE_TYPE_EMF
         * @see Workbook#PICTURE_TYPE_JPEG
         * @see Workbook#PICTURE_TYPE_PICT
         * @see Workbook#PICTURE_TYPE_PNG
         * @see Workbook#PICTURE_TYPE_WMF
         */
        public PictureType PictureType;

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}
