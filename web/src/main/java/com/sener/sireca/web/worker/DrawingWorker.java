/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.DibujoConfTipologia;
import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.service.DibujoServiceImpl;

public class DrawingWorker extends Thread
{
    // Revisi�n de la cual calcular el dibujo de replanteo
    private DibujoRevision revision;
    private DibujoConfTipologia dibConfTip;
    private double pkIni;
    private double pkFin;
    private boolean bHDC;
    private String catenaria;

    public DrawingWorker(DibujoRevision revision,
            DibujoConfTipologia dibConfTip, double pkIni, double pkFin,
            boolean bHDC, String catenaria)
    {
        super();
        this.revision = revision;
        this.dibConfTip = dibConfTip;
        this.pkIni = pkIni;
        this.pkFin = pkFin;
        this.bHDC = bHDC;
        this.catenaria = catenaria;
    }

    @Override
    public void run()
    {

        DibujoServiceImpl service = new DibujoServiceImpl();
        service.calculateRevision(revision, dibConfTip, pkIni, pkFin, bHDC,
                catenaria);

    }
}
