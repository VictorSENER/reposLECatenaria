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
import com.sener.sireca.web.bean.DibujoConfTipologia;
import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
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

            if (parameters[3].equals("E.xlsx"))
                dibujoRevisionAux.setError(true);

            else if (parameters[3].equals("C.xlsx"))
                dibujoRevisionAux.setCalculated(true);

            else if (parameters[3].equals("CW.xlsx"))
            {
                dibujoRevisionAux.setCalculated(true);
                dibujoRevisionAux.setWarning(true);
            }

            if (fileService.fileExists(dibujoRevisionAux.getNotesFilePath()))
                dibujoRevisionAux.setNotes(true);

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
            if (fileService.getFileExtension(ficheros[i]).equals("dwg"))
                revisionList.add(ficheros[i].getName());

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
    public DibujoRevision createRevision(DibujoVersion version,
            ReplanteoRevision repRev, String comment)
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
        lastDibujoRevision.setRepRev(repRev);

        if (!comment.equals(""))
            fileService.writeFile(lastDibujoRevision.getNotesFilePath(),
                    comment);

        return lastDibujoRevision;
    }

    @Override
    public void calculateRevision(DibujoRevision revision,
            DibujoConfTipologia dibConfTip, double pkIni, double pkFin,
            int repVersion, int repRevision)
    {
        JACOBService jacobService = (JACOBService) SpringApplicationContext.getBean("jacobService");

        String path = revision.getBasePath();

        List<Variant> parameter = new ArrayList<Variant>();

        parameter.add(new Variant(dibConfTip.isGeoPost()));
        parameter.add(new Variant(dibConfTip.isEtiPost()));
        parameter.add(new Variant(dibConfTip.isDatPost()));
        parameter.add(new Variant(dibConfTip.isVanos()));
        parameter.add(new Variant(dibConfTip.isFlechas()));
        parameter.add(new Variant(dibConfTip.isDescentramientos()));
        parameter.add(new Variant(dibConfTip.isImplantacion()));
        parameter.add(new Variant(dibConfTip.isAltHilo()));
        parameter.add(new Variant(dibConfTip.isDistCant()));
        parameter.add(new Variant(dibConfTip.isConexiones()));
        parameter.add(new Variant(dibConfTip.isProtecciones()));
        parameter.add(new Variant(dibConfTip.isPendolado()));
        parameter.add(new Variant(dibConfTip.isAltCat()));
        parameter.add(new Variant(dibConfTip.isPuntSing()));
        parameter.add(new Variant(dibConfTip.isCableado()));
        parameter.add(new Variant(dibConfTip.isDatTraz()));

        File preAutoCad = new File(revision.getAutoCadPath());
        File preError = new File(revision.getErrorFilePath());
        File preComment = new File(revision.getNotesFilePath());

        String auxExcelPath = revision.getBasePath() + ".xlsx";

        fileService.fileCopy(revision.getRepRev().getExcelPath(), auxExcelPath);

        if (jacobService.executeCoreCommand(path, "dibujo-replanteo", parameter))
        {
            fileService.deleteFile(revision.getProgressFilePath());
            // fileService.deleteFile(auxExcelPath);
            revision.setCalculated(true);
        }

        if (fileService.fileExists(preError.getAbsolutePath()))
        {
            ArrayList<String[]> errorLog = null;

            try
            {
                errorLog = fileService.getErrorFileContent(path + ".error");

                revision.setWarning(true);
                for (int i = 0; i < errorLog.size(); i++)
                    if (errorLog.get(i)[0].equals("Error"))
                    {
                        revision.setError(true);
                        revision.setWarning(false);
                        break;
                    }
            }
            catch (IOException e)
            {

            }

        }

        revision.changeState(preAutoCad, preError, preComment);

    }

    // Delete the specific revision of the specific version of the specific
    // project and the progress file.
    @Override
    public boolean deleteRevision(Project project, int numVersion,
            int numRevision)
    {
        DibujoVersion version = getVersion(project, numVersion);
        DibujoRevision revision = getRevision(version, numRevision);

        try
        {
            if (!revision.getCalculated())
                return false;
        }
        catch (Exception e)
        {
            return false;
        }

        if (fileService.fileExists(revision.getProgressFilePath()))
            fileService.deleteFile(revision.getProgressFilePath());

        if (fileService.fileExists(revision.getErrorFilePath()))
            fileService.deleteFile(revision.getErrorFilePath());

        if (fileService.fileExists(revision.getNotesFilePath()))
            fileService.deleteFile(revision.getNotesFilePath());

        return fileService.deleteFile(revision.getAutoCadPath());

    }

    @Override
    public String[] getProgressInfo(DibujoRevision revision) throws IOException
    {
        String[] valores = { "0", "?", "Ejecutando funcionalidad desconocida.",
                "0", "?" };

        return fileService.getProgressFileContent(
                revision.getProgressFilePath(), valores);

    }

    @Override
    public ArrayList<String> getNotes(DibujoRevision revision)
            throws IOException
    {
        return fileService.getFileContent(revision.getNotesFilePath());
    }

    @Override
    public ArrayList<String[]> getErrorLog(DibujoRevision revision)
            throws IOException
    {
        return fileService.getErrorFileContent(revision.getErrorFilePath());
    }

}
