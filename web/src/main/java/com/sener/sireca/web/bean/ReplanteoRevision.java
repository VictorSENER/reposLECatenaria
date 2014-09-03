/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.Date;

public class ReplanteoRevision
{
    // Identificador del proyecto al que pertenece la versión.
    private int idProject;

    // Número de versión.
    private int numVersion;

    // Número de revisión.
    private int numRevision;

    // Tipo de revisión (0:calculado, 1:importado).
    private int type;

    // Indica si la revisión ha sido calculada o si aún se está calculando.
    private boolean calculated;

    // Fecha de creación de la revisión.
    private Date creationDate;

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

    public boolean isCalculated()
    {
        return calculated;
    }

    public void setCalculated(boolean calculated)
    {
        this.calculated = calculated;
    }

    public Date getCreationDate()
    {
        return creationDate;
    }

    public void setCreationDate(Date creationDate)
    {
        this.creationDate = creationDate;
    }

    public long getFileSize()
    {
        return fileSize;
    }

    public void setFileSize(long fileSize)
    {
        this.fileSize = fileSize;
    }

    public String getExcelPath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/" + "projects";
        return basePath + "/" + idProject + "/" + "replanteo" + "/"
                + numVersion + "/" + numRevision + "-" + type + ".xls";
    }

    public String getProgressFilePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/" + "projects";
        return basePath + "/" + idProject + "/" + "replanteo" + "/"
                + numVersion + "/" + numRevision + "-" + type + ".txt";
    }
}
