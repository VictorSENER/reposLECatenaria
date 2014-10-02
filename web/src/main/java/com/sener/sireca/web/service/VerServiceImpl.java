/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.util.SpringApplicationContext;

@Service("verService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class VerServiceImpl implements VerService
{

    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

    // Get the list of the version directories and parse it into an Integer
    // ArrayList.
    @Override
    public ArrayList<Integer> getVersions(String ruta)
    {
        ArrayList<Integer> versionList = new ArrayList<Integer>();

        File[] ficheros = fileService.getDirectory(ruta);

        for (int i = 0; i < ficheros.length; i++)
            try
            {
                versionList.add(Integer.parseInt(ficheros[i].getName()));
            }
            catch (Exception e)
            {
                // Ignora el elemento.
            }

        Collections.sort(versionList);

        return versionList;
    }

    // Check if an specific version exists.
    @Override
    public boolean getVersion(String ruta, int version)
    {
        ArrayList<Integer> versionList = getVersions(ruta);

        for (int i = 0; i < versionList.size(); i++)
            if (versionList.get(i) == version)
                return true;

            else if (versionList.get(i) > version)
                return false;

        return false;
    }

    // Get the last version of a project.
    @Override
    public int getLastVersion(String ruta)
    {
        ArrayList<Integer> versionList = getVersions(ruta);
        return versionList.get(versionList.size() - 1);
    }

}
