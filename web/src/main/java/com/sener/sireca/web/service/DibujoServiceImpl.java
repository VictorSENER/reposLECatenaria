/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("dibujoService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class DibujoServiceImpl implements DibujoService
{
    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
    VerService verService = (VerService) SpringApplicationContext.getBean("verService");

    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    // Return a list of the versions of the specific project.
    @Override
    public List<DibujoVersion> getVersions(Project project)
    {
        ArrayList<Integer> versionList = verService.getVersions(project.getDibReplanteoBasePath());
        ArrayList<DibujoVersion> dibujoVersion = new ArrayList<DibujoVersion>();

        for (int i = 0; i < versionList.size(); i++)
            dibujoVersion.add(new DibujoVersion(project.getId(), versionList.get(i)));

        return dibujoVersion;
    }

    @Override
    public List<Integer> getVersionList(Project project)
    {
        return verService.getVersions(project.getDibReplanteoBasePath());
    }

    // Check if the folder exists, and if so build the object.
    @Override
    public DibujoVersion getVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getDibReplanteoBasePath(), numVersion))
            return new DibujoVersion(project.getId(), numVersion);

        return null;
    }

    // Creates a new version of a project.
    @Override
    public DibujoVersion createVersion(Project project)
    {
        int idLastversion = verService.getLastVersion(project.getDibReplanteoBasePath());
        idLastversion++;

        fileService.addDirectory(project.getCalcReplanteoBasePath()
                + idLastversion);

        return new DibujoVersion(project.getId(), idLastversion);
    }

    @Override
    public int getLastVersion(Project project)
    {
        return verService.getLastVersion(project.getDibReplanteoBasePath());
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
    public List<DibujoRevision> getRevisions(DibujoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<DibujoRevision> dibujoRevision = new ArrayList<DibujoRevision>();

        Project project = projectService.getProjectById(version.getIdProject());

        for (int i = 0; i < revisionList.size(); i++)
        {

            String fileName = revisionList.get(i);
            String[] parameters = fileName.split("_");

            DibujoRevision dibujoRevisionAux = new DibujoRevision();
            dibujoRevisionAux.setIdProject(version.getIdProject());
            dibujoRevisionAux.setNumVersion(version.getNumVersion());
            dibujoRevisionAux.setNumRevision(Integer.parseInt(parameters[0]));

            ReplanteoVersion replanteoVersionAux = replanteoService.getVersion(
                    project, Integer.parseInt(parameters[1]));
            dibujoRevisionAux.setRepRev(replanteoService.getRevision(
                    replanteoVersionAux, Integer.parseInt(parameters[2])));

            if (parameters[3].equals("E.dwg"))
                dibujoRevisionAux.setError(true);
            else
                dibujoRevisionAux.setError(false);

            if (parameters[3].equals("C.dwg"))
                dibujoRevisionAux.setCalculated(true);
            else
                dibujoRevisionAux.setCalculated(false);

            dibujoRevisionAux.setDate(fileService.getFileDate(version.getFolderPath()
                    + fileName));
            dibujoRevisionAux.setFileSize(fileService.getFileSize(version.getFolderPath()
                    + fileName));

            dibujoRevision.add(dibujoRevisionAux);

        }

        return dibujoRevision;
    }

    @Override
    public List<Integer> getRevisionList(DibujoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<Integer> revList = new ArrayList<Integer>();

        for (int i = 0; i < revisionList.size(); i++)
        {
            String fileName = revisionList.get(i);
            String[] parameters = fileName.split("_");

            revList.add(Integer.parseInt(parameters[0]));
        }

        return revList;
    }

    // Get the list of the revisions and parse it into a String ArrayList.
    private ArrayList<String> getRevisions(String ruta)
    {
        ArrayList<String> revisionList = new ArrayList<String>();
        File[] ficheros = fileService.getDirectory(ruta);

        for (int i = 0; i < ficheros.length; i++)
        {
            // TODO: Buscar cual va a ser el fichero principal
            if (fileService.getFileExtension(ficheros[i]).equals("dwg"))
                revisionList.add(ficheros[i].getName());
        }

        return revisionList;
    }

    // Returns a specific revision of a specific version.
    @Override
    public DibujoRevision getRevision(DibujoVersion version, int numRevision)
    {
        List<DibujoRevision> dibujoRevision = getRevisions(version);

        for (int i = 0; i < dibujoRevision.size(); i++)
            if (dibujoRevision.get(i).getNumRevision() == numRevision)
                return dibujoRevision.get(i);

        return null;
    }

    // Creates a new revision of the specific version of a project.
    @Override
    public DibujoRevision createRevision(DibujoVersion version, int type)
    {
        int lastRevision = 0;

        List<DibujoRevision> dibujoRevision = getRevisions(version);

        for (int i = 0; i < dibujoRevision.size(); i++)
            if (dibujoRevision.get(i).getNumRevision() > lastRevision)
                lastRevision = dibujoRevision.get(i).getNumRevision();

        DibujoRevision lastDibujoRevision = new DibujoRevision();

        lastDibujoRevision.setIdProject(version.getIdProject());
        lastDibujoRevision.setNumVersion(version.getNumVersion());
        lastDibujoRevision.setNumRevision(lastRevision + 1);

        if (type == 0)
            lastDibujoRevision.setCalculated(false);
        else
            lastDibujoRevision.setCalculated(true);

        lastDibujoRevision.setDate(new Date());

        lastDibujoRevision.setFileSize(fileService.getFileSize(lastDibujoRevision.getAutoCadPath()));

        return lastDibujoRevision;

    }

    @Override
    public void calculateRevision(DibujoRevision revision)
    {
        // TODO: 1) Crea el fichero de progreso: está en blanco
        // 2) Carga los módulos VB sobre el Excel
        // 3) Ejecuta el cálculo VB
    }

    // Delete the specific revision of the specific version of the specific
    // project and the progress file.
    @Override
    public void deleteRevision(Project project, int numVersion, int numRevision)
    {
        DibujoVersion version = getVersion(project, numVersion);
        DibujoRevision revision = getRevision(version, numRevision);

        fileService.deleteFile(revision.getAutoCadPath());
        fileService.deleteFile(revision.getProgressFilePath());

    }

}
