/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.sql.Date;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

@Service("fileService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class FileServiceImpl implements FileService
{

    // Add directory in a specific path.
    public void addDirectory(String ruta)
    {
        File directory = new File(ruta);
        directory.mkdirs();

    }

    // Remove an specific directory.
    public boolean deleteDirectory(String ruta)
    {
        File directory = new File(ruta);
        deleteRecurivelly(directory);

        if (directory.delete())
            return true;

        return false;

    }

    // Remove the specific file.
    public boolean deleteFile(String ruta)
    {
        File file = new File(ruta);

        if (file.delete())
            return true;

        return false;
    }

    // Returns the content of a directory in an array.
    public File[] getDirectory(String ruta)
    {

        File directory = new File(ruta);
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
    public Date getFileDate(String ruta)
    {
        // Comprobar si funciona
        File file = new File(ruta);
        long ms = file.lastModified();

        return new Date(ms);
    }

    // Returns the size of an specific file in bytes.
    public long getFileSize(String ruta)
    {
        // Comprobar si funciona
        File file = new File(ruta);

        return file.length();
    }

}
