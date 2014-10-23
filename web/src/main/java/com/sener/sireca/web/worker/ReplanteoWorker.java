/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.service.ReplanteoServiceImpl;

public class ReplanteoWorker extends Thread
{
    // Revisión de la cual calcular el cuaderno de replanteo
    private ReplanteoRevision revision;
    private String catenaria;
    private double pkIni;
    private double pkFin;

    public ReplanteoWorker(ReplanteoRevision revision, double pkIni,
            double pkFin, String catenaria)
    {
        super();
        this.revision = revision;
        this.catenaria = catenaria;
        this.pkIni = pkIni;
        this.pkFin = pkFin;
    }

    @Override
    public void run()
    {
        ReplanteoServiceImpl service = new ReplanteoServiceImpl();
        service.calculateRevision(this.revision, this.pkIni, this.pkFin,
                this.catenaria);
    }
}
