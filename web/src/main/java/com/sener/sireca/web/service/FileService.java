/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.IOException;
import java.sql.Date;

public interface FileService
{
    public boolean addDirectory(String path);

    public boolean deleteDirectory(String path);

    public boolean deleteFile(String path);

    public boolean addFile(String path);

    public File[] getDirectory(String path);

    public Date getFileDate(String path);

    public long getFileSize(String path);

    public String getFileExtension(File file);

    public void fileCopy(String initPath, String finalPath);

    public String getFileContent(String path) throws IOException;
}
