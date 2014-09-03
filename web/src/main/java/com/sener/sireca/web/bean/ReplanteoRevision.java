/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.Date;

public class ReplanteoRevision
{
    // Identificador del proyecto al que pertenece la versi�n.
    private int idProject;

    // N�mero de versi�n.
    private int numVersion;

    // N�mero de revisi�n.
    private int numRevision;

    // Tipo de revisi�n (0:calculado, 1:importado).
    private int type;

    // Indica si la revisi�n ha sido calculada o si a�n se est� calculando.
    private boolean calculated;

    // Fecha de creaci�n de la revisi�n.
    private Date creationDate;

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
