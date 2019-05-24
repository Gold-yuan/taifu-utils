package sichuan.ytf.pdf.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPRow;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;

import sichuan.ytf.main.util.Base64Util;

public class ExportPDF {

    /**
     * 将图片转换为pdf
     * 
     * @param imgPath Gif, Jpeg, Png
     * @param pdfPath
     */
    public static void imgToPdf(String imgPath, String pdfPath) {
        File pdfFile = new File(pdfPath);
        // 创建一个文档对象
        Document doc = new Document();
        PdfWriter writer = null;
        try {
            // 定义输出文件的位置
            writer = PdfWriter.getInstance(doc, new FileOutputStream(pdfFile));
            // 开启文档
            doc.open();

            // 向文档中加入图片
            // 取得图片~~~图片格式：
            Image jpg1 = Image.getInstance(imgPath); // 原来的图片的路径
            // 获得图片的高度
            float width = jpg1.getWidth();
            // 合理压缩，h>w，按w压缩，否则按w压缩
            // int percent=getPercent(heigth, width);
            // 统一按照宽度压缩
            int percent = getPercent2(width);
            // 设置图片居中显示
            jpg1.setAlignment(Image.MIDDLE);
            // 直接设置图片的大小~~~~~~~第三种解决方案，按固定比例压缩
            // jpg1.scaleAbsolute(210.0f, 297.0f);
            // 按百分比显示图片的比例
            jpg1.scalePercent(percent);// 表示是原来图像的比例;
            // 可设置图像高和宽的比例
            // jpg1.scalePercent(50, 100);
            doc.add(jpg1);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭文档并释放资源
            doc.close();
            writer.flush();
            writer.close();
        }
    }

    /**
     * 合并多个图片、pdf为一个pdf
     * 
     * @param paths     图片、pdf的地址
     * @param mergePath 合成后的地址
     * @return
     */
    public static Integer mergeImgAndPdfToPdf(List<String> paths, String mergePath) {
        File pdfFile = new File(mergePath);
        // 创建一个文档对象
        Document doc = new Document();
        // 定义输出文件的位置
        PdfWriter writer = null;
        try {
            writer = PdfWriter.getInstance(doc, new FileOutputStream(pdfFile));
            // 开启文档
            doc.open();
            PdfContentByte cb = writer.getDirectContent();
            for (String path : paths) {
                if (path.toLowerCase().endsWith(".pdf")) {
                    PdfReader reader = new PdfReader(path);
                    int pageOfCurrentReaderPDF = 0;
                    while (pageOfCurrentReaderPDF < reader.getNumberOfPages()) {
                        doc.newPage();
                        pageOfCurrentReaderPDF++;
                        PdfImportedPage page = writer.getImportedPage(reader, pageOfCurrentReaderPDF);
                        cb.addTemplate(page, 0, 0);
                    }
                } else {
                    // 向文档中加入图片
                    // 取得图片~~~图片格式：
                    Image jpg1 = Image.getInstance(path); // 原来的图片的路径
                    // 获得图片的高度
                    float width = jpg1.getWidth();
                    // 合理压缩，h>w，按w压缩，否则按w压缩
                    // int percent=getPercent(heigth, width);
                    // 统一按照宽度压缩
                    int percent = getPercent2(width);
                    // 设置图片居中显示
                    jpg1.setAlignment(Image.MIDDLE);
                    // 直接设置图片的大小~~~~~~~第三种解决方案，按固定比例压缩
                    // jpg1.scaleAbsolute(210.0f, 297.0f);
                    // 按百分比显示图片的比例
                    jpg1.scalePercent(percent);// 表示是原来图像的比例;
                    // 可设置图像高和宽的比例
                    // jpg1.scalePercent(50, 100);
                    doc.add(jpg1);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } finally {
            // 关闭文档并释放资源
            try {
                if (doc != null && doc.isOpen())
                    doc.close();
                if (writer != null)
                    writer.flush();
                if (writer != null)
                    writer.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return 0;
    }

    public static Integer mergeByPath(List<String> paths, String mergePath) {
        try {
            List<PdfReader> readers = new ArrayList<PdfReader>();
            for (String path : paths) {
                PdfReader reader = new PdfReader(path);
                readers.add(reader);
            }
            return merge(readers, mergePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return 0;
    }

    public static Integer merge(List<PdfReader> readers, String mergePath) {
        if (readers == null || readers.size() == 0) {
            return null;
        }
        int totalPages = 0;
        Document document = new Document(PageSize.A4);
        FileOutputStream out = null;
        PdfWriter writer = null;
        try {
            out = new FileOutputStream(mergePath);
            writer = PdfWriter.getInstance(document, out);
            document.open();
            PdfContentByte cb = writer.getDirectContent();

            int pageOfCurrentReaderPDF = 0;
            Iterator<PdfReader> iteratorPDFReader = readers.iterator();

            // Loop through the PDF files and add to the output.
            while (iteratorPDFReader.hasNext()) {
                PdfReader pdfReader = iteratorPDFReader.next();
                totalPages += pdfReader.getNumberOfPages();
                // Create a new page in the target for each source page.
                while (pageOfCurrentReaderPDF < pdfReader.getNumberOfPages()) {
                    document.newPage();
                    pageOfCurrentReaderPDF++;
                    PdfImportedPage page = writer.getImportedPage(pdfReader, pageOfCurrentReaderPDF);
                    cb.addTemplate(page, 0, 0);
                }
                pageOfCurrentReaderPDF = 0;
            }
            out.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                document.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            try {
                out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            for (PdfReader pdfReader : readers) {
                try {
                    pdfReader.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return totalPages;
    }

    public static String exportPDF(String[] imgStrs, String text) {
        String[] title = { "预测结果查看", "整体总收入预测结果曲线", "整体可供吨公里变化曲线" };
        String[] imgPath = new String[imgStrs.length];
        // 对字节数组字符串进行Base64解码并生成图片
        try {
            for (int j = 0; j < imgStrs.length; j++) {
                String imgStr = imgStrs[j].substring("data:image/png;base64,".length());
                // Base64解码
                byte[] b = Base64Util.decode(imgStr);
                for (int i = 0; i < b.length; ++i) {
                    if (b[i] < 0) {// 调整异常数据
                        b[i] += 256;
                    }
                }
                // 生成jpeg图片
                String imgFilePath = "temp" + System.currentTimeMillis() + ".png";// 新生成的图片
                File file = new File(imgFilePath);
                imgPath[j] = file.getAbsolutePath();
                OutputStream out = new FileOutputStream(file);
                out.write(b);
                out.flush();
                out.close();
                Thread.sleep(10);
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

        File pdfFile = new File("预测结果查看" + System.currentTimeMillis() + ".pdf");
        // 1.新建document对象
        Document document = new Document(PageSize.A4);
        try {
            // 2.建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中。
            // 创建 PdfWriter 对象 第一个参数是对文档对象的引用，第二个参数是文件的实际名称，在该名称中还会给出其输出路径。
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(pdfFile));

            // 3.打开文档
            document.open();
            // 设置中文字体
            BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            // 中文字体
            Font titleFont = new Font(bfChinese, 14, Font.BOLD);
            Font th = new Font(bfChinese);
            th.setSize(10);

            // 添加内容 文字-图片
            for (int i = 0; i < imgPath.length; i++) {

                // 4.添加一个内容段落
                Paragraph titlep = new Paragraph(title[i], titleFont);
                document.add(titlep);

                // 添加一个表格
                if (i == 0) {
                    // 7列的表.
                    PdfPTable table = new PdfPTable(7);

                    table.setWidthPercentage(100); // 宽度100%填充
                    table.setSpacingBefore(10f); // 前间距
                    table.setSpacingAfter(10f); // 后间距

                    List<PdfPRow> listRow = table.getRows();
                    // 设置列宽
                    float[] columnWidths = { 5f, 4f, 4f, 4f, 4f, 4f, 4f };
                    table.setWidths(columnWidths);

                    String[] rowsStr = text.split("\\|");
                    for (int x = 0; x < rowsStr.length; x++) {
                        String[] cellsStr = rowsStr[x].split("!");
                        // 行
                        PdfPCell cells[] = new PdfPCell[cellsStr.length];
                        PdfPRow row = new PdfPRow(cells);
                        for (int z = 0; z < cellsStr.length; z++) {
                            // 单元格
                            cells[z] = new PdfPCell(new Paragraph(cellsStr[z], th));// 单元格内容
                            cells[z].setVerticalAlignment(Element.ALIGN_MIDDLE);// 垂直居中
                        }
                        // 把第一行添加到集合
                        listRow.add(row);
                    }
                    // 把表格添加到文件中
                    document.add(table);
                }

                // 图片
                Image container = Image.getInstance(imgPath[i]);
                container.scalePercent(getPercent2(container.getWidth()));
                container.setAlignment(Element.ALIGN_CENTER);
                document.add(container);
            }
            document.close();
            writer.flush();
            writer.close();
            return pdfFile.getAbsolutePath();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            for (String path : imgPath) {
                File file = new File(path);
                if (file != null && file.isFile() && file.exists()) {
                    file.delete();
                }
            }
        }
        return null;
    }

    /**
     * 第一种解决方案 在不改变图片形状的同时，判断，如果h>w，则按h压缩，否则在w>h或w=h的情况下，按宽度压缩
     * 
     * @param h
     * @param w
     * @return
     */

    public static int getPercent(float h, float w) {
        int p = 0;
        float p2 = 0.0f;
        if (h > w) {
            p2 = 297 / h * 100;
        } else {
            p2 = 210 / w * 100;
        }
        p = Math.round(p2);
        return p;
    }

    /**
     * 第二种解决方案，统一按照宽度压缩 这样来的效果是，所有图片的宽度是相等的，自我认为给客户的效果是最好的
     * 
     * @param args
     */
    public static int getPercent2(float w) {
        int p = 0;
        float p2 = 0.0f;
        p2 = 530 / w * 100;
        p = Math.round(p2);
        return p;
    }
}
