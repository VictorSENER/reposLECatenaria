/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.sener.sireca.web.service.FileService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.UserService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class MontajeRevision
{

    // Identificador del proyecto al que pertenece la versi�n.
    private int idProject;

    // N�mero de versi�n.
    private int numVersion;

    // N�mero de revisi�n.
    private int numRevision;

    // Revisi�n del cuaderno de replanteo usada
    private ReplanteoRevision repRev;

    // Indica si la revisi�n ha sido calculada o si a�n se est� calculando.
    private boolean calculated;

    // Indica si la revisi�n tiene errores fatales o no.
    private boolean error;

    // Indica si la revisi�n tiene warnings o no.
    private boolean warning;

    // Indica si la revisi�n tiene comentarios o no.
    private boolean notes;

    // Fecha de creaci�n de la revisi�n.
    private Date date;

    // Tama�o del fichero de la revisi�n (en bytes).
    private long fileSize;

    public int getIdProject()
    {
        return idProject;
    }

    public void setIdProject(int idProject)
    {
        this.idProject = idProject;
    }

    public int getNumVersion()
    {
        return numVersion;
    }

    public void setNumVersion(int numVersion)
    {
        this.numVersion = numVersion;
    }

    public int getNumRevision()
    {
        return numRevision;
    }

    public void setNumRevision(int numRevision)
    {
        this.numRevision = numRevision;
    }

    public boolean getCalculated()
    {
        return calculated;
    }

    public void setCalculated(boolean calculated)
    {
        this.calculated = calculated;
    }

    public boolean getError()
    {
        return error;
    }

    public void setError(boolean error)
    {
        this.error = error;
    }

    public boolean getWarning()
    {
        return warning;
    }

    public void setWarning(boolean warning)
    {
        this.warning = warning;
    }

    public boolean getNotes()
    {
        return notes;
    }

    public void setNotes(boolean notes)
    {
        this.notes = notes;
    }

    public Date getDate()
    {
        return date;
    }

    public void setDate(Date date)
    {
        this.date = date;
    }

    public long getFileSize()
    {
        return fileSize;
    }

    public void setFileSize(long fileSize)
    {
        this.fileSize = fileSize;
    }

    public ReplanteoRevision getRepRev()
    {
        return repRev;
    }

    public void setRepRev(ReplanteoRevision repRev)
    {
        this.repRev = repRev;
    }

    public void changeState(File preAutoCad, File preError, File preComment,
            File prePDF)
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        File postAutoCad = new File(getAutoCadPath());
        File postPDF = new File(getPDFPath());
        File postError = new File(getErrorFilePath());
        File postComment = new File(getNotesFilePath());

        fileService.rename(prePDF, postPDF);
        fileService.rename(preAutoCad, postAutoCad);
        fileService.rename(preError, postError);
        fileService.rename(preComment, postComment);
    }

    private String getState()
    {
        if (calculated && warning)
            return "_CW";

        else if (error)
            return "_E";

        else if (calculated)
            return "_C";

        else
            return "_P";
    }

    public String getAutoCadPath()
    {
        return getBasePath() + ".dwg";

    }

    public String getAutoCadName()
    {
        return getBaseName() + ".dwg";

    }

    public String getPDFPath()
    {
        return getBasePath() + ".pdf";

    }

    public String getPDFName()
    {
        return getBaseName() + ".pdf";

    }

    public String getProgressFilePath()
    {
        return getBasePath() + ".progress";
    }

    public String getErrorFilePath()
    {
        return getBasePath() + ".error";
    }

    public String getNotesFilePath()
    {
        return getBasePath() + ".comment";
    }

    public String getRUser()
    {
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");

        int idUser = projectService.getProjectById(idProject).getIdUsuario();

        return userService.getUserById(idUser).getUsername();
    }

    public String getRFileSize()
    {
        if (fileSize < 1024)
            return fileSize + " B";

        int exp = (int) (Math.log(fileSize) / Math.log(1024));
        String pre = "" + ("KMGTPE").charAt(exp - 1);

        return String.format("%.1f %sB", fileSize / Math.pow(1024, exp), pre);
    }

    public String getRDate()
    {
        return new SimpleDateFormat("dd-MM-yyyy").format(date);
    }

    public String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + MontajeVersion.FICHAS_MONTAJE
                + numVersion + "/" + getBaseName();
    }

    private String getBaseName()
    {
        return numRevision + "_" + repRev.getNumVersion() + "_"
                + repRev.getNumRevision() + getState();
    }

}
