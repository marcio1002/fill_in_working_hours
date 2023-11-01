<?php

namespace Marcio1002\Work;

class Std
{
    public static function read(): mixed
    {
        return \trim(\fgets(STDIN));
    }

    public static function write(mixed $data): void
    {
        $data = static::format($data);

        \fwrite(STDOUT,  "\e[32m$data\e[m" . PHP_EOL);
    }

    public static function error(mixed $data): void
    {

        $data = static::format($data);

        $exts = \join("|", $GLOBALS['extensions_valid']);

        $data = \preg_replace(
            pattern: "/(\/(?:.)+\.($exts))+/",
            replacement: "\e[33;4m$1\e[m\e[31m",
            subject: $data
        );

        \fwrite(STDERR,  "\e[31m$data\e[m" . PHP_EOL);
    }

    private static function format(mixed $data): string
    {
        if (\is_array($data) && \key_exists(0, $data)) $data = \join("", $data);

        if (\is_array($data) && !\key_exists(0, $data)) $data = \json_encode($data);

        if (!\is_string($data)) $data = (string) $data;

        return $data;
    }
}
