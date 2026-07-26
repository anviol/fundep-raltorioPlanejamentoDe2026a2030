$pdf_mode = 1;
$pdflatex = 'pdflatex -synctex=1 -interaction=nonstopmode -file-line-error %O %S';

# --- Suporte a glossaries-extra / acronyms
# Faz o latexmk chamar makeglossaries automaticamente entre as passadas.
add_cus_dep('acn', 'acr', 0, 'makeglossaries');
add_cus_dep('glo', 'gls', 0, 'makeglossaries');

sub makeglossaries {
    my ($base_name, $path) = fileparse($_[0]);
    my $return = system("makeglossaries", "-d", $path, $base_name);
    return $return;
}

$clean_ext .= ' acn acr alg glg glo gls ist xdy synctex.gz run.xml';
