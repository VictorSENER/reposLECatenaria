/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;

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
        File[] ficheros = directory.listFiles();

        sortDirectory(ficheros);

        return ficheros;
    }

    private void sortDirectory(File[] files)
    {
        Arrays.sort(files, new Comparator<File>()
        {
            public int compare(File o1, File o2)
            {
                int n1 = extractNumber(o1.getName());
                int n2 = extractNumber(o2.getName());
                return n1 - n2;
            }

            private int extractNumber(String name)
            {
                int i = 0;
                try
                {
                    int s = 0;
                    int e = name.indexOf('_');
                    String number = name.substring(s, e);
                    i = Integer.parseInt(number);
                }
                catch (Exception e)
                {
                    i = 0;
                }
                return i;
            }
        });

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

    public String[] getProgressFileContent(String path, String valores[])
            throws IOException
    {

        BufferedReader br = null;

        try
        {
            br = new BufferedReader(new FileReader(path));

            try
            {
                String line = br.readLine();

                if (line != null && line != "")
                    valores = line.split("/");
            }
            finally
            {
                br.close();
            }
        }
        catch (FileNotFoundException e)
        {
            // Ignore
        }

        return valores;

    }

    public ArrayList<String[]> getErrorFileContent(String path)
            throws IOException
    {

        ArrayList<String[]> errorLog = new ArrayList<String[]>();

        BufferedReader br = null;

        try
        {
            br = new BufferedReader(new FileReader(path));

            try
            {

                String line = br.readLine();

                while (line != null && line != "")
                {
                    errorLog.add(line.split("/"));
                    line = br.readLine();
                }
            }
            finally
            {
                br.close();
            }
        }
        catch (FileNotFoundException e)
        {
            // Ignore
        }

        return errorLog;

    }

    public ArrayList<String> getFileContent(String path) throws IOException
    {
        ArrayList<String> notes = new ArrayList<String>();

        BufferedReader br = null;

        try
        {
            br = new BufferedReader(new FileReader(path));

            try
            {

                String line = br.readLine();

                while (line != null)
                {
                    notes.add(line);
                    line = br.readLine();
                }
            }
            finally
            {
                br.close();
            }
        }
        catch (FileNotFoundException e)
        {
            // Ignore
        }

        return notes;
    }

    // public String getFileContent(String path) throws IOException
    // {
    //
    // FileInputStream inputStream = new FileInputStream(path);
    // String everything;
    // try
    // {
    // everything = IOUtils.toString(inputStream);
    // }
    // finally
    // {
    // inputStream.close();
    // }
    //
    // return everything;
    //
    // }

    public boolean fileExists(String path)
    {

        BufferedReader br = null;

        try
        {
            br = new BufferedReader(new FileReader(path));

            try
            {
                br.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }

        }
        catch (FileNotFoundException e)
        {
            return false;
        }

        return true;

    }

    public void rename(File from, File to)
    {
        from.renameTo(to);
    }

    @Override
    public void writeFile(String path, String content)
    {
        PrintWriter writer = null;
        try
        {
            writer = new PrintWriter(path, "UTF-8");
            writer.println(content);
        }
        catch (FileNotFoundException | UnsupportedEncodingException e)
        {
            e.printStackTrace();
        }
        finally
        {
            writer.close();
        }

    }
}
