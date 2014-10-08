/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.text.SimpleDateFormat;
import java.util.Date;

import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.UserService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class DibujoRevision
{

    // Identificador del proyecto al que pertenece la versión.
    private int idProject;

    // Número de versión.
    private int numVersion;

    // Revisión del cuaderno de replanteo usada
    private ReplanteoRevision repRev;

    // Revisión del cuaderno de replanteo usada
    private DibujoConfTipologia confTipo;

    // Número de revisión.
    private int numRevision;

    // Indica si la revisión ha sido calculada o si aún se está calculando.
    private boolean calculated;

    // Indica si la revisión tiene errores o no.
    private boolean error;

    // Fecha de creación de la revisión.
    private Date date;

    // Tamaño del fichero de la revisión (en bytes).
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

    public DibujoConfTipologia getConfTipo()
    {
        return confTipo;
    }

    public void setConfTipo(DibujoConfTipologia confTipo)
    {
        this.confTipo = confTipo;
    }

    private String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + DibujoVersion.DIBUJO_REPLANTEO
                + numVersion + "/" + getBaseName();
    }

    private String getBaseName()
    {
        return numRevision + "_" + repRev.getNumVersion() + "_"
                + repRev.getNumRevision();
    }

    public String getAutoCadPath()
    {
        // TODO: Get extension ".dwg" or ".dwf"
        if (error)
            return getBasePath() + "_E.dwg";

        if (calculated)
            return getBasePath() + "_C.dwg";

        else
            return getBasePath() + "_P.dwg";
    }

    public String getProgressFilePath()
    {
        if (error)
            return getBasePath() + "_E.txt";

        if (calculated)
            return getBasePath() + "_C.txt";

        else
            return getBasePath() + "_P.txt";
    }

    public String getAutoCadName()
    {
        if (error)
            return getBaseName() + "_E.dwg";

        if (calculated)
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
        String pre = "" + ("KMGTPE").charAt(exp - 1); // ("i");

        return String.format("%.1f %sB", fileSize / Math.pow(1024, exp), pre);
    }

    public String getRDate()
    {
        return new SimpleDateFormat("dd-MM-yyyy").format(date);
    }

}
