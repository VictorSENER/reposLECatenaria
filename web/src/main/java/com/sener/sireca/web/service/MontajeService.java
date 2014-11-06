/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;

public interface MontajeService
{
    public List<MontajeVersion> getVersions(Project project);

    public MontajeVersion getVersion(Project project, int numVersion);

    public List<Integer> getVersionList(Project project);

    public MontajeVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<MontajeRevision> getRevisions(MontajeVersion version);

    public List<Integer> getRevisionList(MontajeVersion version);

    public MontajeRevision getRevision(MontajeVersion version, int numRevision);

    public MontajeRevision createRevision(MontajeVersion version,
            ReplanteoRevision repRev, String comment);

    public void calculateRevision(MontajeRevision revision, double pkIni,
            double pkFin, String catenaria, boolean pdf, boolean cad);

    public void deleteRevision(Project project, int numVersion, int numRevision)
            throws Exception;

    public int getLastVersion(Project project);

    String[] getProgressInfo(MontajeRevision revision) throws IOException;

    ArrayList<String[]> getErrorLog(MontajeRevision revision)
            throws IOException;

    ArrayList<String> getNotes(MontajeRevision revision) throws IOException;

    public int getLastRevision(MontajeVersion version);

    public boolean hasMontajeDependencies(Project project, int numVersion,
            int numRevision);

    public List<String> getTemplatesList(Project project);

}
