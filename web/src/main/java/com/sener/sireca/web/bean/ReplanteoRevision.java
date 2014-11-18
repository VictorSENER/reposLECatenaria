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
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@SuppressWarnings("rawtypes")
public class ReplanteoRevision implements Comparable
{

    // Identificador del proyecto al que pertenece la versión.
    private Integer idProject;

    // Número de versión.
    private int numVersion;

    // Número de revisión.
    private int numRevision;

    // Tipo de revisión (0:calculado, 1:importado, 2:recalculado a partir de
    // fase 4)
    private int type;

    // Indica si la revisión ha sido calculada o si aún se está calculando.
    private boolean calculated;

    // Indica si la revisión tiene errores fatales o no.
    private boolean error;

    // Indica si la revisión tiene warnings o no.
    private boolean warning;

    // Indica si la revisión tiene comentarios o no.
    private boolean notes;

    // Fecha de creación de la revisión.
    private Date date;

    // Tamaño del fichero de la revisión (en bytes).
    private long fileSize;

    public Integer getIdProject()
    {
        return idProject;
    }

    public void setIdProject(Integer idProject)
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

    public int getType()
    {
        return type;
    }

    public void setType(int type)
    {
        this.type = type;
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

    @Override
    public int compareTo(Object arg0)
    {
        int compareRev = ((ReplanteoRevision) arg0).getNumRevision();

        return this.numRevision - compareRev;
    }

    public void changeState(File preExcel, File preError, File preComment)
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

        File postExcel = new File(getExcelPath());
        File postError = new File(getErrorFilePath());
        File postComment = new File(getNotesFilePath());

        fileService.rename(preExcel, postExcel);
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

    public String getExcelPath()
    {
        return getBasePath() + ".xlsx";
    }

    public String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            basePath += "/projects/";
        else
            basePath += "/projectTest/";

        return basePath + idProject + ReplanteoVersion.CALCULO_REPLANTEO
                + numVersion + "/" + getBaseName();
    }

    private String getBaseName()
    {

        return numRevision + "_" + type + getState();
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

    public String getRType()
    {
        if (type == 0)
            return "Calculado";

        else if (type == 1)
            return "Importado";

        else
            return "Recalculado";
    }

    public String getRFileSize()
    {
        if (fileSize < 1024)
            return fileSize + " B";

        int exp = (int) (Math.log(fileSize) / Math.log(1024));
        String pre = "" + ("KMGTPE").charAt(exp - 1); // ("i");

        return String.format("%.1f %sB", fileSize / Math.pow(1024, exp), pre);
    }

    public String getRDate()
    {
        return new SimpleDateFormat("dd-MM-yyyy").format(date);
    }

}
