/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.service.PendoladoServiceImpl;

public class PendoladoWorker extends Thread
{
    // Revisión de la cual calcular el cuaderno de replanteo
    private PendoladoRevision revision;
    private String catenaria;
    private double pkIni;
    private double pkFin;

    public PendoladoWorker(PendoladoRevision revision, double pkIni,
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
        PendoladoServiceImpl service = new PendoladoServiceImpl();
        service.calculateRevision(this.revision, this.pkIni, this.pkFin,
                this.catenaria);
    }
}
