/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.DibujoConfTipologia;
import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.service.DibujoServiceImpl;

public class DrawingWorker extends Thread
{
    // Revisión de la cual calcular el dibujo de replanteo
    private DibujoRevision revision;
    private DibujoConfTipologia dibConfTip;
    private double pkIni;
    private double pkFin;
    private int repVersion;
    private int repRevision;

    public DrawingWorker(DibujoRevision revision,
            DibujoConfTipologia dibConfTip, double pkIni, double pkFin,
            int repVersion, int repRevision)
    {
        super();
        this.revision = revision;
        this.dibConfTip = dibConfTip;
        this.pkIni = pkIni;
        this.pkFin = pkFin;
        this.repVersion = repVersion;
        this.repRevision = repRevision;
    }

    @Override
    public void run()
    {

        DibujoServiceImpl service = new DibujoServiceImpl();
        service.calculateRevision(revision, dibConfTip, pkIni, pkFin,
                repVersion, repRevision);

    }
}
