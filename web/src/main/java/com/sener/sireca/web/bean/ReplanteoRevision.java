/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.text.SimpleDateFormat;
import java.util.Date;

import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.UserService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ReplanteoRevision
{

    // Identificador del proyecto al que pertenece la versión.
    private int idProject;

    // Número de versión.
    private int numVersion;

    // Número de revisión.
    private int numRevision;

    // Tipo de revisión (0:calculado, 1:importado, 2:recalculado a partir de
    // fase 4)
    private int type;

    // Indica si la revisión ha sido calculada o si aún se está calculando.
    private boolean calculated;

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

    private String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + Globals.CALCULO_REPLANTEO + numVersion
                + "/" + numRevision + "_" + type;
    }

    public String getExcelPath()
    {
        if (calculated)
            return getBasePath() + "_C.xlsx";

        else
            return getBasePath() + "_P.xlsx";
    }

    public String getProgressFilePath()
    {
        if (calculated)
            return getBasePath() + "_C.txt";

        else
            return getBasePath() + "_P.txt";
    }

    public String getExcelName()
    {
        if (calculated)
            return numRevision + "_" + type + "_C.xlsx";

        else
            return numRevision + "_" + type + "_P.xlsx";

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
