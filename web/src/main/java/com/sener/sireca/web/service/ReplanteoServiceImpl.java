/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.ArrayList;
import java.util.List;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;

public class ReplanteoServiceImpl
{
    public List<ReplanteoVersion> getVersions(Project project)
    {
        // TODO: mirar carpetas bajo el proyecto
        return new ArrayList<ReplanteoVersion>();
    }

    public ReplanteoVersion getVersion(Project project, int numVersion)
    {
        // TODO: mira si existe carpeta, y si es así construye el objeto
        return null;
    }

    public ReplanteoVersion createVersion(Project project)
    {
        // TODO: calcula número de version (el mayor + 1) y crea una carpeta con
        // el
        return null;
    }

    public void deleteVersion(Project project, int numVersion)
    {
        // TODO: hace primero un getVersion(), y luego borra la carpeta con el
        // path del objeto obtenido.
    }

    public List<ReplanteoRevision> getRevisions(ReplanteoVersion version)
    {
        // TODO: - mirar ficheros bajo la carpeta de la revisión dada.
        // - construir objetos a partir del nombre de los ficheros
        // - considerar tb ficheros de progreso (si existen, calculated = false)
        return new ArrayList<ReplanteoRevision>();
    }

    public ReplanteoRevision getRevision(ReplanteoVersion version,
            int numRevision)
    {
        // TODO: - hace un getRevisions() e itera sobre la lista obtenida para
        // buscar
        // la revisión con el número dado.
        return null;
    }

    public ReplanteoRevision createRevision(ReplanteoVersion version, int type)
    {
        // TODO: calcula el siguiente numero de revision y crea el objeto
        // correspondiente
        // pero no hace nada con los ficheros: la GUI ya se encargará de ponerlo
        // en el path que indica el fichero.
        return null;
    }

    public void calculateRevision(ReplanteoRevision revision)
    {
        // TODO: 1) Crea el fichero de progreso: está en blanco
        // 2) Carga los módulos VB sobre el Excel
        // 3) Ejecuta el cálculo VB
    }

    public void deleteRevision(Project project, int numVersion, int numRevision)
    {
        // TODO: hace primero un getRevision(), y luego borra el fichero con el
        // path del objeto obtenido.
        // tb borra el fichero de progreso si existiera
    }
}
