/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;

import org.apache.commons.io.IOUtils;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

@Service("fileService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class FileServiceImpl implements FileService
{

    // Add directory in a specific path.
    @Override
    public boolean addDirectory(String path)
    {
        File directory = new File(path);
        return directory.mkdirs();

    }

    // Remove an specific directory.
    @Override
    public boolean deleteDirectory(String path)
    {
        File directory = new File(path);
        deleteRecurivelly(directory);

        return directory.delete();
    }

    // Remove the specific file.
    @Override
    public boolean deleteFile(String path)
    {
        File file = new File(path);

        return file.delete();
    }

    // Add a file in the specific Path.
    @Override
    public boolean addFile(String path)
    {
        File file = new File(path);

        try
        {
            file.createNewFile();
            return true;
        }
        catch (IOException e)
        {
            return false;
        }
    }

    // Returns the content of a directory in an array.
    @Override
    public File[] getDirectory(String path)
    {

        File directory = new File(path);
        return directory.listFiles();

    }

    // Delete all the files in a directory.
    private void deleteRecurivelly(File directory)
    {
        File[] ficheros = directory.listFiles();

        for (int x = 0; x < ficheros.length; x++)
        {
            if (ficheros[x].isDirectory())
                deleteRecurivelly(ficheros[x]);

            ficheros[x].delete();
        }
    }

    // Returns the date of an specific file.
    @Override
    public Date getFileDate(String path)
    {
        // Comprobar si funciona
        File file = new File(path);
        long ms = file.lastModified();

        return new Date(ms);
    }

    // Returns the size of an specific file in bytes.
    @Override
    public long getFileSize(String path)
    {
        File file = new File(path);

        return file.length();
    }

    @Override
    public String getFileExtension(File file)
    {
        String fileName = file.getName();
        if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
            return fileName.substring(fileName.lastIndexOf(".") + 1);
        else
            return "";
    }

    @Override
    public void fileCopy(String initPath, String finalPath)
    {
        try
        {
            FileInputStream in = new FileInputStream(initPath);
            BufferedInputStream reader = new BufferedInputStream(in, 4096);
            FileOutputStream out = new FileOutputStream(finalPath);
            BufferedOutputStream writer = new BufferedOutputStream(out, 4096);
            byte[] buf = new byte[4096];
            int byteRead;
            while ((byteRead = reader.read(buf, 0, 4096)) >= 0)
            {
                writer.write(buf, 0, byteRead);
            }
            reader.close();
            writer.flush();
            writer.close();
        }
        catch (Throwable exception)
        {
            exception.printStackTrace();
        }
    }

    @Override
    public String getFileContent(String path) throws IOException
    {

        FileInputStream inputStream = new FileInputStream(path);
        String everything;
        try
        {
            everything = IOUtils.toString(inputStream);
        }
        finally
        {
            inputStream.close();
        }

        return everything;

    }
}
