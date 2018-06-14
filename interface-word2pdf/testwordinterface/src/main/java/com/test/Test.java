package com.test;

import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;
import com.sun.rowset.internal.InsertRow;
import javafx.scene.chart.PieChart;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.List;


public class Test {

  public static Document oDoc;
  double _imgH;
  double _imgW;

  /**
   * 获取注册文件
   */
  public static void getLicense() {
    InputStream is;
    try {

      is=WordUtil.class.getClassLoader().getResourceAsStream("license.xml");
      License license = new License();
      license.setLicense(is);
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }


  public Map<String,Object> resourcesTest()
  {

    Map<String,Object> JBQK=new HashMap<String, Object>();
    JBQK.put("DYH","2018918293891283912");
    JBQK.put("DJJG","晋江");
    JBQK.put("IMG_1","E://303.jpg");


    return JBQK;
  }

  public Map<String,Object> resourcesTest2()
  {
    Map<String,Object> mapMaxBig=new HashMap<String, Object>();

    List<Map<String,Object>> BHQK=new ArrayList<Map<String, Object>>();

    Map<String,Object> row1=new HashMap<String, Object>();
    row1.put("BHYY","变化原因");
    row1.put("BHNR","变化内容是什么呢");
    row1.put("DJSJ","2017-01-01 12:30:45");
    row1.put("DBR","甘建峰");

    Map<String,Object> row2=new HashMap<String, Object>();
    row2.put("BHYY","变化原因");
    row2.put("BHNR","变化内容是什么呢");
    row2.put("DJSJ","2017-01-01 12:30:45");
    row2.put("DBR","甘建峰");

    Map<String,Object> row3=new HashMap<String, Object>();
    row3.put("BHYY","变化原因");
    row3.put("BHNR","变化内容是什么呢");
    row3.put("DJSJ","2017-01-01 12:30:45");
    row3.put("DBR","甘建峰");

    BHQK.add(row1);
    BHQK.add(row2);
    BHQK.add(row3);
    mapMaxBig.put("BHQK",BHQK);

    List<Map<String,Object>> ZRZK=new ArrayList<Map<String, Object>>();

    Map<String,Object> zRow1=new HashMap<String, Object>();
    zRow1.put("LX",1);
    zRow1.put("LB","类别");
    zRow1.put("MJ",120);

    Map<String,Object> zRow2=new HashMap<String, Object>();
    zRow2.put("LX",1);
    zRow2.put("LB","类别");
    zRow2.put("MJ",120);

    Map<String,Object> zRow3=new HashMap<String, Object>();
    zRow3.put("LX",1);
    zRow3.put("LB","类别");
    zRow3.put("MJ",120);

    ZRZK.add(zRow1);
    ZRZK.add(zRow2);
    ZRZK.add(zRow3);
    mapMaxBig.put("ZRZK",ZRZK);



    return mapMaxBig;
  }

  public void readMap(Map<String,Object> resource,Map<String,Object> resource2,String fileName) throws Exception {
    OpenWithTemplate(fileName);
    DataSet ds=new DataSet();
    if (!resource.isEmpty())
    {
      DataTable table=new DataTable("JBQK");
      for (Map.Entry<String,Object> tableentry:resource.entrySet())
      {
        table.getColumns().add(tableentry.getKey());
      }
      DataRow dr=table.newRow();
      for (Map.Entry<String,Object> tableentry:resource.entrySet())
      {
        dr.set(tableentry.getKey(),tableentry.getValue());
      }
      table.getRows().add(dr);
      oDoc.getMailMerge().execute(table);
      InsertImg(oDoc,table);
    }
    if (!resource2.isEmpty())
    {
      for (Map.Entry<String,Object> tableentry:resource2.entrySet())
      {
        DataTable table=new DataTable(tableentry.getKey());
        System.out.println(tableentry.getKey());
        ArrayList<Map<String,Object>> rows= (ArrayList<Map<String, Object>>) tableentry.getValue();
        Iterator iterator=rows.iterator();
        while (iterator.hasNext())
        {
          Map<String,Object> map= (Map<String, Object>) iterator.next();
          for (Map.Entry<String,Object> row:map.entrySet())
          {
            table.getColumns().add(row.getKey());
          }
          DataRow dr=table.newRow();
          for (Map.Entry<String,Object> row:map.entrySet())
          {
            dr.set(row.getKey(),row.getValue());
          }
          table.getRows().add(dr);

        }
        //System.out.println(table.toString());
        ds.getTables().add(table);
        oDoc.getMailMerge().executeWithRegions(ds);
      }
    }


    oDoc.save("E://自然资源登记簿NEW.docx");
    oDoc.save("E://自然资源登记簿NEW.pdf",SaveFormat.PDF);

  }

  public String getNewFileName(String fileName)
  {
    String[] newFile=fileName.split(".");
    for (int i=0;i<newFile.length;i++)
    {
      System.out.println(newFile[i].toString());
    }
    return null;
  }

  public static void OpenWithTemplate(String strFileName)
  {
    if (!strFileName.trim().equals("")&&strFileName.length()>0)
    {
      try {
        oDoc = new Document(strFileName);
      } catch (Exception e) {
        e.printStackTrace();
      }
    }
  }

  public void InsertImg(Document doc, DataTable dt) throws Exception {

    String imgPath = "";
    for (int i = 0; i < doc.getRange().getBookmarks().getCount() ; i++)
    {

      Bookmark mark = doc.getRange().getBookmarks().get(i);
      DocumentBuilder builder = new DocumentBuilder(doc);

      if (mark.getName().substring(0,4).equals("IMG_"))
      {

        imgPath=(String) dt.getRows().get(0).get(mark.getName());
        System.out.println(imgPath);
        if (imgPath != "")
        {
          File imgFile=new File(imgPath);
          if (imgFile.exists())
          {
            getImgSize(imgPath);
            builder.moveToBookmark(mark.getName());
            builder.insertImage(imgPath, RelativeHorizontalPosition.MARGIN,10,RelativeHorizontalPosition.MARGIN,1,_imgW,_imgH, WrapType.SQUARE);
          }
        }
      }
    }
  }

  public void getImgSize(String img) {

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
          cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
        }
        else
        {
          cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
        }
        tmp = cellval;
      }

    }
  }

  public static void main(String[] args) {
   getLicense();
    Test test=new Test();
    Map<String,Object> map=test.resourcesTest();
    Map<String,Object> map2=test.resourcesTest2();
    try {
     // test.readMap(map);
      test.readMap(map,map2,"E://自然资源登记簿.docx");
    } catch (Exception e) {
      e.printStackTrace();
    }

  }
}
