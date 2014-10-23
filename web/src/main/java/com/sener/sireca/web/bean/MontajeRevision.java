/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.text.SimpleDateFormat;
import java.util.Date;

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

    // Indica si la revisi�n ha sido calculada o si a�n se est� calculando.
    private boolean calculated;

    // Indica si la revisi�n tiene errores fatales o no.
    private boolean error;

    // Indica si la revisi�n tiene warnings o no.
    private boolean warning;

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

    public String getAutoCadPath()
    {

        if (calculated && warning)
            return getBasePath() + "_CW.dwg";

        else if (error)
            return getBasePath() + "_E.dwg";

        else if (calculated)
            return getBasePath() + "_C.dwg";

        else
            return getBasePath() + "_P.dwg";
    }

    public String getProgressFilePath()
    {
        if (!calculated)
            return getBasePath() + ".progress";

        return "";
    }

    public String getErrorFilePath()
    {
        if (error || warning)
            return getBasePath() + ".error";

        return "";
    }

    public String getAutoCadName()
    {
        if (calculated && warning)
            return getBaseName() + "_CW.dwg";

        else if (error)
            return getBaseName() + "_E.dwg";

        else if (calculated)
            return getBaseName() + "_C.dwg";

        else
            return getBaseName() + "_P.dwg";
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

    private String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + MontajeVersion.FICHAS_MONTAJE
                + numVersion + "/" + getBaseName();
    }

    private String getBaseName()
    {
        // TODO: Completar BaseName
        return "";
    }

}
