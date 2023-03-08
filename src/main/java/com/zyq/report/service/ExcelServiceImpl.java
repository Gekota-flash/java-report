package com.zyq.report.service;

import com.zyq.report.util.LocalDateUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private LocalDateUtil localDateUtil;

    @Override
    public File copyExcel(String path) {
        File source = new File(path);
        System.out.println("path:" + path);
        System.out.println("path lengh:" + path.length());
        System.out.println("isFile:" + source.isFile());
        String newPath = path.substring(0, path.lastIndexOf(".")) + localDateUtil.getDateTimeNow() + ".xlsx";
        File dest = new File(newPath);
        System.out.println("new path:" + newPath);
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
            os.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                os.close();
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return dest;
    }

}
