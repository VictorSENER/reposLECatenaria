/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.service.DibujoServiceImpl;

public class DrawingWorker extends Thread
{
    // Revisión de la cual calcular el dibujo de replanteo
    private DibujoRevision revision;

    // private String catenaria;
    // private long pkIni;
    // private long pkFin;

    public DrawingWorker(DibujoRevision revision)
    {
        super();
        this.revision = revision;
        // this.catenaria = catenaria;
        // this.pkIni = pkIni;
        // this.pkFin = pkFin;
    }

    @Override
    public void run()
    {

        DibujoServiceImpl service = new DibujoServiceImpl();
        service.calculateRevision(revision);

    }
}
