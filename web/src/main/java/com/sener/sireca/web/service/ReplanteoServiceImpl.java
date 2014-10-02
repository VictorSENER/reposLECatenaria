/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.jacob.com.Variant;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("replanteoService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class ReplanteoServiceImpl implements ReplanteoService
{
    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
    VerService verService = (VerService) SpringApplicationContext.getBean("verService");
    JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");

    // Return a list of the versions of the specific project.
    @Override
    public List<ReplanteoVersion> getVersions(Project project)
    {
        ArrayList<Integer> versionList = verService.getVersions(project.getCalcReplanteoBasePath());
        ArrayList<ReplanteoVersion> replanteoVersion = new ArrayList<ReplanteoVersion>();

        for (int i = 0; i < versionList.size(); i++)
            replanteoVersion.add(new ReplanteoVersion(project.getId(), versionList.get(i)));

        return replanteoVersion;
    }

    // Check if the folder exists, and if so build the object.
    @Override
    public ReplanteoVersion getVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getCalcReplanteoBasePath(),
                numVersion))
            return new ReplanteoVersion(project.getId(), numVersion);

        return null;
    }

    // Creates a new version of a project.
    @Override
    public ReplanteoVersion createVersion(Project project)
    {
        int idLastversion = verService.getLastVersion(project.getCalcReplanteoBasePath());
        idLastversion++;

        fileService.addDirectory(project.getCalcReplanteoBasePath()
                + idLastversion);

        return new ReplanteoVersion(project.getId(), idLastversion);
    }

    @Override
    public int getLastVersion(Project project)
    {
        return verService.getLastVersion(project.getCalcReplanteoBasePath());
    }

    // Delete the specific version of a specific project.
    @Override
    public void deleteVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getCalcReplanteoBasePath(),
                numVersion))
            fileService.deleteDirectory(project.getCalcReplanteoBasePath()
                    + numVersion);
    }

    // Return a list of the revisions of a specific project.
    @Override
    public List<ReplanteoRevision> getRevisions(ReplanteoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<ReplanteoRevision> replanteoRevision = new ArrayList<ReplanteoRevision>();

        for (int i = 0; i < revisionList.size(); i++)
        {

            String fileName = revisionList.get(i);
            String[] parameters = fileName.split("_");

            ReplanteoRevision replanteoRevisionAux = new ReplanteoRevision();
            replanteoRevisionAux.setIdProject(version.getIdProject());
            replanteoRevisionAux.setNumVersion(version.getNumVersion());
            replanteoRevisionAux.setNumRevision(Integer.parseInt(parameters[0]));
            replanteoRevisionAux.setType(Integer.parseInt(parameters[1]));
            if (parameters[2].equals("E.xlsx"))
                replanteoRevisionAux.setError(true);

            else if (parameters[2].equals("C.xlsx"))
                replanteoRevisionAux.setCalculated(true);

            else
                replanteoRevisionAux.setCalculated(false);

            replanteoRevisionAux.setDate(fileService.getFileDate(version.getFolderPath()
                    + fileName));
            replanteoRevisionAux.setFileSize(fileService.getFileSize(version.getFolderPath()
                    + fileName));

            replanteoRevision.add(replanteoRevisionAux);

        }

        return replanteoRevision;

    }

    // Get the list of the revisions and parse it into a String ArrayList.
    private ArrayList<String> getRevisions(String ruta)
    {
        ArrayList<String> revisionList = new ArrayList<String>();
        File[] ficheros = fileService.getDirectory(ruta);

        for (int i = 0; i < ficheros.length; i++)
        {
            if (fileService.getFileExtension(ficheros[i]).equals("xlsx"))
                revisionList.add(ficheros[i].getName());
        }

        return revisionList;
    }

    // Returns a specific revision of a specific version.
    @Override
    public ReplanteoRevision getRevision(ReplanteoVersion version,
            int numRevision)
    {
        List<ReplanteoRevision> replanteoRevision = getRevisions(version);

        for (int i = 0; i < replanteoRevision.size(); i++)
            if (replanteoRevision.get(i).getNumRevision() == numRevision)
                return replanteoRevision.get(i);

        return null;
    }

    @Override
    public int getLastRevision(ReplanteoVersion version)
    {
        int lastRevision = 0;

        List<ReplanteoRevision> replanteoRevision = getRevisions(version);

        for (int i = 0; i < replanteoRevision.size(); i++)
            if (replanteoRevision.get(i).getNumRevision() > lastRevision)
                lastRevision = replanteoRevision.get(i).getNumRevision();

        return lastRevision;
    }

    // Creates a new revision of the specific version of a project.
    @Override
    public ReplanteoRevision createRevision(ReplanteoVersion version, int type)
    {

        int lastRevision = getLastRevision(version);

        ReplanteoRevision lastReplanteoRevision = new ReplanteoRevision();

        lastReplanteoRevision.setIdProject(version.getIdProject());
        lastReplanteoRevision.setNumVersion(version.getNumVersion());
        lastReplanteoRevision.setNumRevision(lastRevision + 1);
        lastReplanteoRevision.setType(type);
        if (type == 0)
            lastReplanteoRevision.setCalculated(false);
        else
            lastReplanteoRevision.setCalculated(true);

        lastReplanteoRevision.setDate(new Date());
        lastReplanteoRevision.setFileSize(fileService.getFileSize(lastReplanteoRevision.getExcelPath()));

        return lastReplanteoRevision;

    }

    @Override
    public void calculateRevision(ReplanteoRevision revision, int idCatenaria,
            int pkIni, int pkFin)
    {

        String path = revision.getBasePath();

        List<Variant> parameter = new ArrayList<Variant>();
        parameter.add(new Variant(idCatenaria));
        parameter.add(new Variant(pkIni));
        parameter.add(new Variant(pkFin));

        File file = new File(revision.getExcelPath());

        if (jacobService.executeCoreCommand(path + "_P", "calculo-replanteo",
                parameter))
            revision.setCalculated(true);
        else
            revision.setError(true);

        File newFile = new File(revision.getExcelPath());

        file.renameTo(newFile);

        fileService.deleteFile(revision.getFilePath());
    }

    // Delete the specific revision of the specific version of the specific
    // project and the progress file.
    @Override
    public void deleteRevision(Project project, int numVersion, int numRevision)
    {
        ReplanteoVersion version = getVersion(project, numVersion);
        ReplanteoRevision revision = getRevision(version, numRevision);

        fileService.deleteFile(revision.getExcelPath());
        fileService.deleteFile(revision.getFilePath());

    }

    @Override
    public String[] getProgressInfo(ReplanteoRevision replanteoRevision)
            throws IOException
    {
        String valores[] = { "0", "?" };

        BufferedReader br = null;

        try
        {
            br = new BufferedReader(new FileReader(replanteoRevision.getFilePath()));

            try
            {
                String line = br.readLine();

                if (line != null)
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

    @Override
    public String getErrorLog(ReplanteoRevision replanteoRevision)
            throws IOException
    {
        return fileService.getFileContent(replanteoRevision.getFilePath());

    }
}
