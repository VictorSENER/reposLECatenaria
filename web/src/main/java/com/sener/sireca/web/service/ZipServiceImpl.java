/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

@Service("zipService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class ZipServiceImpl implements ZipService

{

    private List<String> fileList;

    @Override
    public void generateZip(String path)
    {
        fileList = new ArrayList<String>();
        File sf = new File(path);
        generateFileList(path, sf);

        byte[] buffer = new byte[1024];
        FileOutputStream fos = null;
        ZipOutputStream zos = null;
        try
        {

            fos = new FileOutputStream(sf.getAbsolutePath() + ".zip");
            zos = new ZipOutputStream(fos);
            FileInputStream in = null;

            for (String file : this.fileList)
            {
                ZipEntry ze = new ZipEntry(file);
                zos.putNextEntry(ze);
                try
                {
                    in = new FileInputStream(sf.getAbsolutePath()
                            + File.separator + file);
                    int len;
                    while ((len = in.read(buffer)) > 0)
                        zos.write(buffer, 0, len);

                }
                finally
                {
                    in.close();
                }
            }

            zos.closeEntry();

        }
        catch (IOException ex)
        {
            ex.printStackTrace();
        }
        finally
        {
            try
            {
                zos.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }

    private void generateFileList(String path, File node)
    {

        // Add file only
        if (node.isFile())
            fileList.add(generateZipEntry(path, node.toString()));

        // Add file inside directory recursivelly
        if (node.isDirectory())
        {
            String[] subNote = node.list();

            for (String filename : subNote)
                generateFileList(path, new File(node, filename));
        }
    }

    private static String generateZipEntry(String path, String file)
    {
        return file.substring(path.length() + 1, file.length());
    }

}