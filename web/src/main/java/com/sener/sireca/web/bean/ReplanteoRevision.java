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

    public boolean isCalculated()
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

    public String getBasePath()
    {

        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + Globals.CALCULO_REPLANTEO + "/"
                + numVersion + "/" + numRevision + "_" + type;

    }

    public String getExcelPath()
    {

        if (calculated)
            return getBasePath() + "_C.xls";

        else
            return getBasePath() + "_P.xls";

    }

    public String getProgressFilePath()
    {

        if (calculated)
            return getBasePath() + "_C.txt";

        else
            return getBasePath() + "_P.txt";

    }
}
