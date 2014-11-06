/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.sener.sireca.web.bean.DibujoConfTipologia;
import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;

public interface DibujoService
{
    public List<DibujoVersion> getVersions(Project project);

    public DibujoVersion getVersion(Project project, int numVersion);

    public List<Integer> getVersionList(Project project);

    public DibujoVersion createVersion(Project project);

    public void deleteVersion(Project project, int numVersion);

    public List<DibujoRevision> getRevisions(DibujoVersion version);

    public List<Integer> getRevisionList(DibujoVersion version);

    public DibujoRevision getRevision(DibujoVersion version, int numRevision);

    public DibujoRevision createRevision(DibujoVersion version,
            ReplanteoRevision repRev, String comment);

    public void calculateRevision(DibujoRevision revision,
            DibujoConfTipologia dibConfTip, double pkIni, double pkFin,
            boolean bHDC, String catenaria);

    public void deleteRevision(Project project, int numVersion, int numRevision)
            throws Exception;

    public int getLastVersion(Project project);

    String[] getProgressInfo(DibujoRevision revision) throws IOException;

    ArrayList<String[]> getErrorLog(DibujoRevision revision) throws IOException;

    ArrayList<String> getNotes(DibujoRevision revision) throws IOException;

    public int getLastRevision(DibujoVersion version);

    public boolean hasDibujoDependencies(Project project, int numVersion,
            int numRevision);

}
