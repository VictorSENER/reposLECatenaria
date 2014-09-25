/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.sql.Date;

public interface FileService
{
    public void addDirectory(String ruta);

    public boolean deleteDirectory(String ruta);

    public boolean deleteFile(String ruta);

    public boolean addFile(String ruta);

    public File[] getDirectory(String ruta);

    public Date getFileDate(String ruta);

    public long getFileSize(String ruta);

    public String getFileExtension(File file);
}
