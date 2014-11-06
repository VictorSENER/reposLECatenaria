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
import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("montajeService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class MontajeServiceImpl implements MontajeService
{
    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
    VerService verService = (VerService) SpringApplicationContext.getBean("verService");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    // Return a list of the versions of the specific project.
    @Override
    public List<MontajeVersion> getVersions(Project project)
    {
        ArrayList<Integer> versionList = verService.getVersions(project.getMonReplanteoBasePath());
        ArrayList<MontajeVersion> montajeVersion = new ArrayList<MontajeVersion>();

        for (int i = 0; i < versionList.size(); i++)
            if (i + 1 == versionList.size() && i != 0)
                montajeVersion.add(new MontajeVersion(project.getId(), versionList.get(i), true));
            else
                montajeVersion.add(new MontajeVersion(project.getId(), versionList.get(i), false));

        return montajeVersion;
    }

    @Override
    public List<Integer> getVersionList(Project project)
    {
        return verService.getVersions(project.getMonReplanteoBasePath());
    }

    // Check if the folder exists, and if so build the object.
    @Override
    public MontajeVersion getVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getMonReplanteoBasePath(), numVersion))
            return new MontajeVersion(project.getId(), numVersion, false);

        return null;
    }

    // Creates a new version of a project.
    @Override
    public MontajeVersion createVersion(Project project)
    {
        int idLastversion = verService.getLastVersion(project.getMonReplanteoBasePath());
        idLastversion++;

        fileService.addDirectory(project.getMonReplanteoBasePath()
                + idLastversion);

        return new MontajeVersion(project.getId(), idLastversion, true);
    }

    @Override
    public int getLastVersion(Project project)
    {
        return verService.getLastVersion(project.getMonReplanteoBasePath());
    }

    // Delete the specific version of a specific project.
    @Override
    public void deleteVersion(Project project, int numVersion)
    {
        if (verService.getVersion(project.getMonReplanteoBasePath(), numVersion))
            fileService.deleteDirectory(project.getMonReplanteoBasePath()
                    + numVersion);
    }

    // Return a list of the revisions of a specific project.
    @Override
    public List<MontajeRevision> getRevisions(MontajeVersion version)
    {
        ArrayList<String> revisionList = getRevisions(version.getFolderPath());
        ArrayList<MontajeRevision> montajeRevision = new ArrayList<MontajeRevision>();

        Project project = projectService.getProjectById(version.getIdProject());

        for (int i = 0; i < revisionList.size(); i++)
        {

            String fileName = revisionList.get(i);
            String[] parameters = fileName.split("_");

            MontajeRevision montajeRevisionAux = new MontajeRevision();
            montajeRevisionAux.setIdProject(version.getIdProject());
            montajeRevisionAux.setNumVersion(version.getNumVersion());
            montajeRevisionAux.setNumRevision(Integer.parseInt(parameters[0]));

            ReplanteoVersion replanteoVersionAux = replanteoService.getVersion(
                    project, Integer.parseInt(parameters[1]));
            if (replanteoVersionAux != null)
            {
                ReplanteoRevision replanteoRevAux = replanteoService.getRevision(
                        replanteoVersionAux, Integer.parseInt(parameters[2]));
                if (replanteoRevAux != null)
                {
                    montajeRevisionAux.setRepRev(replanteoRevAux);

                    // Son carpetas y archivos .zip
                    if (parameters[3].equals("E.zip")
                            || parameters[3].equals("E"))
                        montajeRevisionAux.setError(true);

                    else if (parameters[3].equals("C.zip"))
                        montajeRevisionAux.setCalculated(true);

                    else if (parameters[3].equals("CW.zip"))
                    {
                        montajeRevisionAux.setCalculated(true);
                        montajeRevisionAux.setWarning(true);
                    }
                    else if (parameters[3].equals("P"))
                        montajeRevisionAux.setCalculated(false);

                    else
                        break;

                    if (fileService.fileExists(montajeRevisionAux.getNotesFilePath()))
                        montajeRevisionAux.setNotes(true);

                    montajeRevisionAux.setDate(fileService.getFileDate(version.getFolderPath()
                            + fileName));
                    montajeRevisionAux.setFileSize(fileService.getFileSize(version.getFolderPath()
                            + fileName));

                    montajeRevision.add(montajeRevisionAux);
                }
            }
        }

        return montajeRevision;
    }

    @Override
    public List<Integer> getRevisionList(MontajeVersion version)
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
            revisionList.add(ficheros[i].getName());

        return revisionList;
    }

    // Returns a specific revision of a specific version.
    @Override
    public MontajeRevision getRevision(MontajeVersion version, int numRevision)
    {
        List<MontajeRevision> montajeRevision = getRevisions(version);

        for (int i = 0; i < montajeRevision.size(); i++)
            if (montajeRevision.get(i).getNumRevision() == numRevision)
                return montajeRevision.get(i);

        return null;
    }

    @Override
    public int getLastRevision(MontajeVersion version)
    {
        int lastRevision = 0;

        List<MontajeRevision> montajeRevision = getRevisions(version);

        for (int i = 0; i < montajeRevision.size(); i++)
            if (montajeRevision.get(i).getNumRevision() > lastRevision)
                lastRevision = montajeRevision.get(i).getNumRevision();

        return lastRevision;
    }

    // Creates a new revision of the specific version of a project.
    @Override
    public MontajeRevision createRevision(MontajeVersion version,
            ReplanteoRevision repRev, String comment)
    {
        int lastRevision = getLastRevision(version);

        List<MontajeRevision> montajeRevision = getRevisions(version);

        for (int i = 0; i < montajeRevision.size(); i++)
            if (montajeRevision.get(i).getNumRevision() > lastRevision)
                lastRevision = montajeRevision.get(i).getNumRevision();

        MontajeRevision lastMontajeRevision = new MontajeRevision();

        lastMontajeRevision.setIdProject(version.getIdProject());
        lastMontajeRevision.setNumVersion(version.getNumVersion());
        lastMontajeRevision.setNumRevision(lastRevision + 1);
        lastMontajeRevision.setRepRev(repRev);

        if (!comment.equals(""))
            fileService.writeFile(lastMontajeRevision.getNotesFilePath(),
                    comment);

        return lastMontajeRevision;
    }

    @Override
    public void calculateRevision(MontajeRevision revision, double pkIni,
            double pkFin, String catenaria, boolean pdf, boolean cad)
    {

        JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");
        ZipService zipService = (ZipService) SpringApplicationContext.getBean("zipService");

        Project project = projectService.getProjectById(revision.getIdProject());

        String path = project.getTemplate(MontajeVersion.FICHAS_MONTAJE);

        fileService.fileCopy(path + "-A_D.dwg", revision.getBasePath()
                + "-A_D.dwg");
        fileService.fileCopy(path + "-A_I.dwg", revision.getBasePath()
                + "-A_I.dwg");
        fileService.fileCopy(path + "-M_D.dwg", revision.getBasePath()
                + "-M_D.dwg");
        fileService.fileCopy(path + "-M_I.dwg", revision.getBasePath()
                + "-M_I.dwg");

        String auxExcelPath = revision.getBasePath() + ".xlsx";

        fileService.fileCopy(revision.getRepRev().getExcelPath(), auxExcelPath);

        path = revision.getBasePath();

        List<Variant> parameter = new ArrayList<Variant>();

        parameter.add(new Variant(pkIni));
        parameter.add(new Variant(pkFin));
        parameter.add(new Variant(catenaria));
        parameter.add(new Variant(pdf));
        parameter.add(new Variant(cad));

        File preFolder = new File(revision.getFolderPath());
        File preZip = new File(revision.getZipPath());
        File preError = new File(revision.getErrorFilePath());
        File preComment = new File(revision.getNotesFilePath());

        if (jacobService.executeCoreCommand(path, "fichas-montaje", parameter))
        {
            if (cad)
                fileService.deleteFile(revision.getBasePath() + ".bak");

            fileService.deleteFile(revision.getProgressFilePath());
            fileService.deleteFile(revision.getBasePath() + "-A_D.dwg");
            fileService.deleteFile(revision.getBasePath() + "-A_I.dwg");
            fileService.deleteFile(revision.getBasePath() + "-M_D.dwg");
            fileService.deleteFile(revision.getBasePath() + "-M_I.dwg");
            fileService.deleteFile(revision.getBasePath() + ".dwg");
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
        MontajeVersion version = getVersion(project, numVersion);

        if (version == null)
            throw new Exception("Error: No se ha podido eliminar la revisión "
                    + numRevision + " de la versión " + numVersion + ".");

        MontajeRevision revision = getRevision(version, numRevision);

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
    public String[] getProgressInfo(MontajeRevision revision)
            throws IOException
    {
        String[] valores = { "0", "?", "Ejecutando funcionalidad desconocida.",
                "0", "?" };

        return fileService.getProgressFileContent(
                revision.getProgressFilePath(), valores);

    }

    @Override
    public ArrayList<String> getNotes(MontajeRevision revision)
            throws IOException
    {
        return fileService.getFileContent(revision.getNotesFilePath());
    }

    @Override
    public ArrayList<String[]> getErrorLog(MontajeRevision revision)
            throws IOException
    {
        return fileService.getErrorFileContent(revision.getErrorFilePath());
    }

    public boolean hasMontajeDependencies(Project project, int numVersion,
            int numRevision)
    {
        List<MontajeVersion> montajeVersion = getVersions(project);

        for (int i = 0; i < montajeVersion.size(); i++)
        {
            List<MontajeRevision> montajeRevision = getRevisions(montajeVersion.get(i));

            for (int j = 0; j < montajeRevision.size(); j++)
                if (montajeRevision.get(i).getRepRev().getNumVersion() == numVersion
                        && montajeRevision.get(i).getRepRev().getNumRevision() == numRevision)
                    return true;

        }
        return false;
    }

    public List<String> getTemplatesList(Project project)
    {
        ArrayList<String> templatesList = new ArrayList<String>();

        File[] templates = fileService.getDirectory(project.getTemplatePath(MontajeVersion.FICHAS_MONTAJE));

        for (int i = 0; i < templates.length; i++)
            if (templates[i].getName().split("-")[1].equals("M_D.dwg"))
                templatesList.add(templates[i].getName().split("-")[0]);

        return templatesList;

    }

}
