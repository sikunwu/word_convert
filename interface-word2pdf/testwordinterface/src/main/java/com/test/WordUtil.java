package com.test;

import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/*import freemarker.template.Configuration;
import freemarker.template.Template;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;*/

public class WordUtil {

  public static Document oDoc;
  double _imgH;
  double _imgW;
  /**
   * 获取注册文件
   */
  public static void getLicense() {
    //String path = getWebRootAbsolutePath() + "/license.xml";
    InputStream is;
    try {
      //is = new FileInputStream(new File(path));
      is=WordUtil.class.getClassLoader().getResourceAsStream("license.xml");
      License license = new License();
      license.setLicense(is);
    } catch (FileNotFoundException e) {
      //logger.error("license.xml file not found");
    } catch (Exception e) {
     // logger.error("license register failed");
    }
  }

  public static void OpenWithTemplate(String strFileName)
  {
    if (strFileName.length()>0&&!strFileName.trim().equals(""))
    {
      try {
        oDoc = new Document(strFileName);
      } catch (Exception e) {
        e.printStackTrace();
      }
    }
  }

  public void InsertImg(Document doc, DataTable dt) throws Exception {
    //foreach (Bookmark mark in doc.Range.Bookmarks)
    //System.out.println("进来了");
    String imgPath = "";
    for (int i = 0; i < doc.getRange().getBookmarks().getCount() ; i++)
    {
     // System.out.println("进来了2");
      Bookmark mark = doc.getRange().getBookmarks().get(i);
      DocumentBuilder builder = new DocumentBuilder(doc);
      //System.out.println(mark.getName().substring(0,4));
      if (mark.getName().substring(0,4).equals("IMG_"))
      {
        //System.out.println("进来了3");
        //imgPath = Convert.ToString(dt.Rows[0][mark.Name]);
        imgPath=(String) dt.getRows().get(0).get(mark.getName());
        System.out.println(imgPath);
        if (imgPath != "")
        {
          File imgFile=new File(imgPath);
          if (imgFile.exists())
          {
            getImgSize(imgPath);
           // builder.MoveToBookmark(mark.Name);
            builder.moveToBookmark(mark.getName());
           // builder.InsertImage(imgPath, RelativeHorizontalPosition.Margin, 10, RelativeVerticalPosition.Margin, 1, _imgW, _imgH, WrapType.Square);
            builder.insertImage(imgPath,RelativeHorizontalPosition.MARGIN,10,RelativeHorizontalPosition.MARGIN,1,_imgW,_imgH,WrapType.SQUARE);
          }
        }
      }
    }
  }

  public void getImgSize(String img) {
   /* double rate;
    using (FileStream fs = new FileStream(img, FileMode.Open, FileAccess.Read))
    {
      System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
      _imgW = image.Width;
      _imgH = image.Height;
      rate = _imgH / _imgW;
      if (_imgW > 500)
      {
        _imgW = 500;
        _imgH = 500 * rate;
      }
      if (_imgH > 830)
      {
        _imgH = 830;
        _imgW = 830 / rate;
      }

    }*/
    try {
      BufferedImage image= ImageIO.read(new File(img));
      _imgW=image.getWidth();
      _imgH=image.getHeight();
      double rate=_imgH/_imgW;
      if (_imgW > 500)
      {
        _imgW = 500;
        _imgH = 500 * rate;
      }
      if (_imgH > 830)
      {
        _imgH = 830;
        _imgW = 830 / rate;
      }

    } catch (IOException e) {
      e.printStackTrace();
    }




  }

  public void VerticalCellMerge(Document doc, int tableIndex, int rowIndex, int colIndex, int rowNum) throws Exception {
    NodeCollection tables =doc.getChildNodes(NodeType.TABLE,true);
    Table table = (Table) tables.get (tableIndex);
    String tmp = "", cellval = "";

    //foreach (Table table in tables)
    //{
    // Get the index of the table node as contained in the parent node of the table
    //int tableIndex = table.ParentNode.ChildNodes.IndexOf(table);
    //Console.WriteLine("Start of Table {0}", tableIndex);
    // Iterate through all rows in the table

    for (Row row:table.getRows())
    {
      int rIndex =table.getRows().indexOf(row);
      if (rIndex > rowIndex + rowNum)
        break;
      else if (rIndex > rowIndex)
      {
        Cell cell =row.getCells().get(colIndex);
        cellval = cell.toString(SaveFormat.TEXT).trim();

        if (tmp == "")
        {
          tmp = cellval;
        }
        if (tmp == cellval)
        {
          //cell.CellFormat.VerticalMerge = CellMerge.Previous;
          cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
        }
        else
        {
          //cell.CellFormat.VerticalMerge = CellMerge.First;
          cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
        }
        //cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        tmp = cellval;
      }

    }
  }

  public void test() throws Exception {
    String oldFile = "E://自然资源登记簿.docx";
    String newFile = "E://自然资源登记簿New.docx";
    String pdf = "E://自然资源登记簿New.pdf";
    String img = "E://210.jpg";
    OpenWithTemplate(oldFile);

    DataTable jbqk=new DataTable("JBQK");
    jbqk.getColumns().add("DYH");
    jbqk.getColumns().add("DJJG");
    jbqk.getColumns().add("IMG_1");
    DataRow dr1=jbqk.newRow();
    dr1.set("DYH","2018918293891283912");
    dr1.set("DJJG","晋江");
    dr1.set("IMG_1","E://303.jpg");
    jbqk.getRows().add(dr1);

    DataTable table=new DataTable("BHQK");

    table.getColumns().add("BHYY");

    table.getColumns().add("BHNR");

    table.getColumns().add("DJSJ");

    table.getColumns().add("DBR");

    for (int i=0;i<10;i++)
    {
      DataRow dr=table.newRow();
      dr.set("BHYY","变化原因" + i);
      dr.set("BHNR","变化内容是什么呢");
      dr.set("DJSJ","2017-01-01 12:30:45");
      dr.set("DBR","甘建峰");
      table.getRows().add(dr);
    }

    DataTable zrzk=new DataTable("ZRZK");
    zrzk.getColumns().add("LX");
    zrzk.getColumns().add("LB");
    zrzk.getColumns().add("MJ");

    for (int i=0;i<8;i++)
    {
      String lx="1";
      if (i>5)
        lx="0";
      if (i>8)
        lx="5";
      if(i>15)
        lx="6";
      DataRow dr=zrzk.newRow();
      dr.set("LX",lx);
      dr.set("LB","类别" + i);
      dr.set("MJ","120");

      zrzk.getRows().add(dr);
    }

    //合并模版，相当于页面的渲染
    DataSet ds=new DataSet();
    ds.getTables().add(table);
    ds.getTables().add(zrzk);
    oDoc.getMailMerge().executeWithRegions(ds);
    oDoc.getMailMerge().execute(jbqk);
    //InsertImg(oDoc,jbqk);
    VerticalCellMerge(oDoc,0,6,1,ds.getTables().get("ZRZK").getRows().getCount());
    oDoc.save(newFile);
    oDoc.save(pdf,SaveFormat.PDF);

  }

  public static void main(String[] args) throws Exception {
    getLicense();
   /* LoadOptions loadOptions = new LoadOptions();
    loadOptions.setLoadFormat(com.aspose.words.LoadFormat.HTML);
    Document doc = new Document("E://hero.html", loadOptions);
    doc.save("E://hreo.doc");*/
    WordUtil wordUtil=new WordUtil();
    wordUtil.test();





  }
}
