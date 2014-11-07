/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.jacob.com.Variant;
import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("pendoladoService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class PendoladoServiceImpl implements PendoladoService
{
    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
    VerService verService = (VerService) SpringApplicationContext.getBean("verService");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    // Return a list of the versions of the specific project.
    @Override
    public List<PendoladoVersion> getVersions(Project project)
    {
        ArrayList<Integer> versionList = verService.getVersions(project.getPenReplanteoBasePath());
        ArrayList<PendoladoVersion> pendoladoVersion = new ArrayList<PendoladoVersion>();

        for (int i = 0; i < versionList.size(); i++)
            if (i + 1 == versionList.size() && i != 0)
                pendoladoVersion.add(new PendoladoVersion(project.getId(), versionList.get(i), true));
            else
                pendoladoVersion.add(new PendoladoVersion(project.getId(), versionList.get(i), false));

        return pendoladoVersion;
    }

    @Override
    public List<Integer> getVersionList(Project project)
    {
        return verService.getVersions(project.getPenReplanteoBasePath());
    }

    // Check if the folder exists, and if so build the object.
    @Override
    public PendoladoVersion getVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getPenReplanteoBasePath(), numVersion))
            return new PendoladoVersion(project.getId(), numVersion, false);

        return null;
    }

    // Creates a new version of a project.
    @Override
    public PendoladoVersion createVersion(Project project)
    {
        int idLastversion = verService.getLastVersion(project.getPenReplanteoBasePath());
        idLastversion++;

        fileService.addDirectory(project.getPenReplanteoBasePath()
                + idLastversion);

        return new PendoladoVersion(project.getId(), idLastversion, true);
    }

    @Override
    public int getLastVersion(Project project)
    {
        return verService.getLastVersion(project.getPenReplanteoBasePath());
    }

    // Delete the specific version of a specific project.
    @Override
    public void deleteVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getPenReplanteoBasePath(), numVersion))
            fileService.deleteDirectory(project.getPenReplanteoBasePath()
                    + numVersion);
    }

    // Return a list of the revisions of a specific project.
    @Override
    public List<PendoladoRevision> getRevisions(PendoladoVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<PendoladoRevision> pendoladoRevision = new ArrayList<PendoladoRevision>();

        Project project = projectService.getProjectById(version.getIdProject());
        for (int i = 0; i < revisionList.size(); i++)
        {
            try
            {
                String fileName = revisionList.get(i);
                String[] parameters = fileName.split("_");

                PendoladoRevision pendoladoRevisionAux = new PendoladoRevision();
                pendoladoRevisionAux.setIdProject(version.getIdProject());
                pendoladoRevisionAux.setNumVersion(version.getNumVersion());
                pendoladoRevisionAux.setNumRevision(Integer.parseInt(parameters[0]));

                ReplanteoVersion replanteoVersionAux = replanteoService.getVersion(
                        project, Integer.parseInt(parameters[1]));

                if (replanteoVersionAux != null)
                {
                    ReplanteoRevision replanteoRevAux = replanteoService.getRevision(
                            replanteoVersionAux,
                            Integer.parseInt(parameters[2]));
                    if (replanteoRevAux != null)
                    {

                        pendoladoRevisionAux.setRepRev(replanteoRevAux);

                        if (parameters[3].equals("E.zip")
                                || parameters[3].equals("E"))
                            pendoladoRevisionAux.setError(true);

                        else if (parameters[3].equals("C.zip"))
                            pendoladoRevisionAux.setCalculated(true);

                        else if (parameters[3].equals("CW.zip"))
                        {
                            pendoladoRevisionAux.setCalculated(true);
                            pendoladoRevisionAux.setWarning(true);
                        }
                        else if (parameters[3].equals("P"))
                            pendoladoRevisionAux.setCalculated(false);

                        else
                            continue;

                        if (fileService.fileExists(pendoladoRevisionAux.getNotesFilePath()))
                            pendoladoRevisionAux.setNotes(true);

                        pendoladoRevisionAux.setDate(fileService.getFileDate(version.getFolderPath()
                                + fileName));
                        pendoladoRevisionAux.setFileSize(fileService.getFileSize(version.getFolderPath()
                                + fileName));

                        pendoladoRevision.add(pendoladoRevisionAux);
                    }
                }
            }
            catch (Exception e)
            {
            }
        }
        return pendoladoRevision;
    }

    @Override
    public List<Integer> getRevisionList(PendoladoVersion version)
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
            if (ficheros[i].isDirectory()
                    || fileService.getFileExtension(ficheros[i]).equals("zip"))
                revisionList.add(ficheros[i].getName());

        return revisionList;
    }

    // Returns a specific revision of a specific version.
    @Override
    public PendoladoRevision getRevision(PendoladoVersion version,
            int numRevision)
    {
        List<PendoladoRevision> pendoladoRevision = getRevisions(version);

        for (int i = 0; i < pendoladoRevision.size(); i++)
            if (pendoladoRevision.get(i).getNumRevision() == numRevision)
                return pendoladoRevision.get(i);

        return null;
    }

    @Override
    public int getLastRevision(PendoladoVersion version)
    {
        int lastRevision = 0;

        List<PendoladoRevision> pendoladoRevision = getRevisions(version);

        for (int i = 0; i < pendoladoRevision.size(); i++)
            if (pendoladoRevision.get(i).getNumRevision() > lastRevision)
                lastRevision = pendoladoRevision.get(i).getNumRevision();

        return lastRevision;
    }

    // Creates a new revision of the specific version of a project.
    @Override
    public PendoladoRevision createRevision(PendoladoVersion version,
            ReplanteoRevision repRev, String comment)
    {
        int lastRevision = getLastRevision(version);

        PendoladoRevision lastPendoladoRevision = new PendoladoRevision();

        lastPendoladoRevision.setIdProject(version.getIdProject());
        lastPendoladoRevision.setNumVersion(version.getNumVersion());
        lastPendoladoRevision.setNumRevision(lastRevision + 1);
        lastPendoladoRevision.setRepRev(repRev);

        if (!comment.equals(""))
            fileService.writeFile(lastPendoladoRevision.getNotesFilePath(),
                    comment);

        return lastPendoladoRevision;
    }

    @Override
    public void calculateRevision(PendoladoRevision revision, double pkIni,
            double pkFin, String catenaria)
    {
        JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");
        ZipService zipService = (ZipService) SpringApplicationContext.getBean("zipService");

        Project project = projectService.getProjectById(revision.getIdProject());

        fileService.addDirectory(revision.getFolderPath());

        String path = project.getTemplate(PendoladoVersion.FICHAS_PENDOLADO);

        fileService.fileCopy(path + "-C.xls", revision.getBasePath() + "-C.xls");
        fileService.fileCopy(path + "-P.xlsm", revision.getBasePath()
                + "-P.xlsm");

        String auxExcelPath = revision.getBasePath() + ".xlsx";

        fileService.fileCopy(revision.getRepRev().getExcelPath(), auxExcelPath);

        path = revision.getBasePath();

        List<Variant> parameter = new ArrayList<Variant>();

        parameter.add(new Variant(pkIni));
        parameter.add(new Variant(pkFin));
        parameter.add(new Variant(catenaria));

        File preFolder = new File(revision.getFolderPath());
        File preZip = new File(revision.getZipPath());
        File preError = new File(revision.getErrorFilePath());
        File preComment = new File(revision.getNotesFilePath());

        if (jacobService.executeCoreCommand(path, "fichas-pendolado", parameter))
        {
            fileService.deleteFile(revision.getProgressFilePath());
            fileService.deleteFile(revision.getBasePath() + "-P.xlsm");
            fileService.deleteFile(revision.getBasePath() + "-C.xls");
            fileService.deleteFile(auxExcelPath);
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
                errorLog = fileService.getErrorFileContent(preError.getName());

                for (int i = 0; i < errorLog.size(); i++)
                    if (errorLog.get(i)[0].equals("Error"))
                    {
                        revision.setError(true);
                        revision.setWarning(false);
                        break;
                    }
                    else
                        revision.setWarning(true);
            }
            catch (IOException e)
            {

            }

        }

        zipService.generateZip(preFolder.getAbsolutePath());

        revision.changeState(preZip, preError, preComment);

        fileService.deleteDirectory(preFolder.getAbsolutePath());

        if (revision.getError() != true && revision.getWarning() != true)
            fileService.deleteFile(revision.getErrorFilePath());

    }

    // Delete the specific revision of the specific version of the specific
    // project and the progress file.
    @Override
    public void deleteRevision(Project project, int numVersion, int numRevision)
            throws Exception
    {
        PendoladoVersion version = getVersion(project, numVersion);

        if (version == null)
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

        PendoladoRevision revision = getRevision(version, numRevision);

        if (revision == null || !revision.getCalculated())
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

        if (fileService.fileExists(revision.getProgressFilePath()))
            fileService.deleteFile(revision.getProgressFilePath());

        if (fileService.fileExists(revision.getErrorFilePath()))
            fileService.deleteFile(revision.getErrorFilePath());

        if (fileService.fileExists(revision.getNotesFilePath()))
            fileService.deleteFile(revision.getNotesFilePath());

        if (!fileService.deleteFile(revision.getZipPath()))
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

    }

    @Override
    public String[] getProgressInfo(PendoladoRevision revision)
            throws IOException
    {
        String[] valores = { "0", "?", "Ejecutando funcionalidad desconocida.",
                "0", "?" };

        return fileService.getProgressFileContent(
                revision.getProgressFilePath(), valores);

    }

    @Override
    public ArrayList<String> getNotes(PendoladoRevision revision)
            throws IOException
    {
        return fileService.getFileContent(revision.getNotesFilePath());
    }

    @Override
    public ArrayList<String[]> getErrorLog(PendoladoRevision revision)
            throws IOException
    {
        return fileService.getErrorFileContent(revision.getErrorFilePath());
    }

    public boolean hasPendoladoDependencies(Project project, int numVersion,
            int numRevision)
    {

        List<PendoladoVersion> pendoladoVersion = getVersions(project);

        for (int i = 0; i < pendoladoVersion.size(); i++)
        {
            List<PendoladoRevision> pendoladoRevision = getRevisions(pendoladoVersion.get(i));

            for (int j = 0; j < pendoladoRevision.size(); j++)
                if (pendoladoRevision.get(i).getRepRev().getNumVersion() == numVersion
                        && pendoladoRevision.get(i).getRepRev().getNumRevision() == numRevision)
                    return true;

        }
        return false;
    }

    public List<String> getTemplatesList(Project project)
    {
        ArrayList<String> templatesList = new ArrayList<String>();

        File[] templates = fileService.getDirectory(project.getTemplatePath(PendoladoVersion.FICHAS_PENDOLADO));

        for (int i = 0; i < templates.length; i++)
            if (templates[i].getName().split("-")[1].equals("P.xlsm"))
                templatesList.add(templates[i].getName().split("-")[0]);

        return templatesList;

    }
}
