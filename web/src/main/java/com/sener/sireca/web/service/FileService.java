/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.IOException;
import java.sql.Date;
import java.util.ArrayList;

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

    public String[] getProgressFileContent(String path) throws IOException;

    public ArrayList<String> getErrorFileContent(String path)
            throws IOException;

    public String getFileContent(String path) throws IOException;

    public boolean fileExists(String path);

}
