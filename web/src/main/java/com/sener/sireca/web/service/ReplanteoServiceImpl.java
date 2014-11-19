/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
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

    // Return a list of the versions of the specific project.
    @Override
    public List<ReplanteoVersion> getVersions(Project project)
    {
        ArrayList<Integer> versionList = verService.getVersions(project.getCalcReplanteoBasePath());
        ArrayList<ReplanteoVersion> replanteoVersion = new ArrayList<ReplanteoVersion>();

        for (int i = 0; i < versionList.size(); i++)
            if (i == versionList.size() - 1 && i != 0)
                replanteoVersion.add(new ReplanteoVersion(project.getId(), versionList.get(i), true));
            else
                replanteoVersion.add(new ReplanteoVersion(project.getId(), versionList.get(i), false));

        return replanteoVersion;
    }

    @Override
    public List<Integer> getVersionList(Project project)
    {
        return verService.getVersions(project.getCalcReplanteoBasePath());
    }

    // Check if the folder exists, and if so build the object.
    @Override
    public ReplanteoVersion getVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getCalcReplanteoBasePath(),
                numVersion))
            return new ReplanteoVersion(project.getId(), numVersion, false);

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

        return new ReplanteoVersion(project.getId(), idLastversion, true);
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
    @SuppressWarnings("unchecked")
    @Override
    public List<ReplanteoRevision> getRevisions(ReplanteoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<ReplanteoRevision> replanteoRevision = new ArrayList<ReplanteoRevision>();

        for (int i = 0; i < revisionList.size(); i++)
        {
            try
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

                else if (parameters[2].equals("CW.xlsx"))
                {
                    replanteoRevisionAux.setCalculated(true);
                    replanteoRevisionAux.setWarning(true);
                }

                if (fileService.fileExists(replanteoRevisionAux.getNotesFilePath()))
                    replanteoRevisionAux.setNotes(true);

                replanteoRevisionAux.setDate(fileService.getFileDate(version.getFolderPath()
                        + fileName));
                replanteoRevisionAux.setFileSize(fileService.getFileSize(version.getFolderPath()
                        + fileName));

                replanteoRevision.add(replanteoRevisionAux);

            }
            catch (Exception e)
            {
            }

        }

        Collections.sort(replanteoRevision);

        return replanteoRevision;

    }

    public List<Integer> getRevisionList(ReplanteoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<Integer> revList = new ArrayList<Integer>();

        for (int i = 0; i < revisionList.size(); i++)
        {
            String fileName = revisionList.get(i);
            String[] parameters = fileName.split("_");

            if (parameters[2].equals("C.xlsx")
                    || parameters[2].equals("CW.xlsx"))
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
    public ReplanteoRevision createRevision(ReplanteoVersion version, int type,
            String comment)
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

        if (!comment.equals(""))
            fileService.writeFile(lastReplanteoRevision.getNotesFilePath(),
                    comment);

        return lastReplanteoRevision;

    }

    @Override
    public void calculateRevision(ReplanteoRevision revision, double pkIni,
            double pkFin, String catenaria)
    {

        JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");

        String path = revision.getBasePath();

        List<Variant> parameter = new ArrayList<Variant>();

        parameter.add(new Variant(pkIni));
        parameter.add(new Variant(pkFin));
        parameter.add(new Variant(catenaria));

        File preExcel = new File(revision.getExcelPath());
        File preError = new File(revision.getErrorFilePath());
        File preComment = new File(revision.getNotesFilePath());

        if (jacobService.executeCoreCommand(path, "calculo-replanteo",
                parameter))
        {
            fileService.deleteFile(revision.getProgressFilePath());
            revision.setCalculated(true);
        }
        else
        {
            fileService.writeFile(preError.getAbsolutePath(),
                    "ERROR/Error en la ejecución del Core.");
            revision.setError(true);
        }

        if (fileService.fileExists(preError.getAbsolutePath()))
        {
            ArrayList<String[]> errorLog = null;

            try
            {
                errorLog = fileService.getErrorFileContent(preError.getAbsolutePath());

                for (int i = 0; i < errorLog.size(); i++)
                    if (errorLog.get(i)[0].equals("Error"))
                    {
                        revision.setError(true);
                        revision.setWarning(false);
                        i = errorLog.size();
                    }
                    else
                        revision.setWarning(true);
            }
            catch (IOException e)
            {
            }

        }

        revision.changeState(preExcel, preError, preComment);

        if (revision.getError() != true && revision.getWarning() != true)
            fileService.deleteFile(revision.getErrorFilePath());

    }

    // Delete the specific revision of the specific version of the specific
    // project and the progress file.
    @Override
    public void deleteRevision(Project project, int numVersion, int numRevision)
            throws Exception
    {
        ReplanteoVersion version = getVersion(project, numVersion);

        if (version == null)
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

        ReplanteoRevision revision = getRevision(version, numRevision);
        if (revision == null || !revision.getCalculated())
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

        // Check if revision has dependencies.
        DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
        if (dibujoService.hasDibujoDependencies(project, numVersion,
                numRevision))
            throw new Exception("Error: La revisión " + numRevision
                    + " de la versión " + numVersion
                    + " tiene dependencias con Planos de replanteo.");

        PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
        if (pendoladoService.hasPendoladoDependencies(project, numVersion,
                numRevision))
            throw new Exception("Error: La revisión " + numRevision
                    + " de la versión " + numVersion
                    + " tiene dependencias con Fichas de pendolado.");

        MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
        if (montajeService.hasMontajeDependencies(project, numVersion,
                numRevision))
            throw new Exception("Error: La revisión " + numRevision
                    + " de la versión " + numVersion
                    + " tiene dependencias con Fichas de montaje.");

        // Delete revision
        if (fileService.fileExists(revision.getErrorFilePath()))
            fileService.deleteFile(revision.getErrorFilePath());

        if (fileService.fileExists(revision.getProgressFilePath()))
            fileService.deleteFile(revision.getProgressFilePath());

        if (fileService.fileExists(revision.getNotesFilePath()))
            fileService.deleteFile(revision.getNotesFilePath());

        if (!fileService.deleteFile(revision.getExcelPath()))
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");
    }

    @Override
    public String[] getProgressInfo(ReplanteoRevision revision)
            throws IOException
    {
        String[] valores = { "0", "?", "Ejecutando funcionalidad desconocida.",
                "0", "?" };

        return fileService.getProgressFileContent(
                revision.getProgressFilePath(), valores);
    }

    @Override
    public ArrayList<String[]> getErrorLog(ReplanteoRevision revision)
            throws IOException
    {
        return fileService.getErrorFileContent(revision.getErrorFilePath());
    }

    public ArrayList<String> getNotes(ReplanteoRevision revision)
            throws IOException
    {
        return fileService.getFileContent(revision.getNotesFilePath());
    }
}
