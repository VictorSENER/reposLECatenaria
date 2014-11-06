/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;

public interface PendoladoService
{
    public List<PendoladoVersion> getVersions(Project project);

    public PendoladoVersion getVersion(Project project, int numVersion);

    public List<Integer> getVersionList(Project project);

    public PendoladoVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<PendoladoRevision> getRevisions(PendoladoVersion version);

    public List<Integer> getRevisionList(PendoladoVersion version);

    public PendoladoRevision getRevision(PendoladoVersion version,
            int numRevision);

    public PendoladoRevision createRevision(PendoladoVersion version,
            ReplanteoRevision repRev, String comment);

    public void calculateRevision(PendoladoRevision revision, double pkIni,
            double pkFin, String catenaria);

    public void deleteRevision(Project project, int numVersion, int numRevision)
            throws Exception;

    public int getLastVersion(Project project);

    String[] getProgressInfo(PendoladoRevision revision) throws IOException;

    ArrayList<String[]> getErrorLog(PendoladoRevision revision)
            throws IOException;

    ArrayList<String> getNotes(PendoladoRevision revision) throws IOException;

    public int getLastRevision(PendoladoVersion version);

    public boolean hasPendoladoDependencies(Project project, int numVersion,
            int numRevision);

    public List<String> getTemplatesList(Project project);

}
