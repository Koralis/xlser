<?php

namespace Sakalys\Xlser;

class XlsWriter
{
    /**
     * @var Xlser
     */
    private $manager;

    /**
     * XlsWriter constructor.
     * @param Xlser $manager
     */
    public function __construct($manager)
    {
        $this->manager = $manager;
    }
}
