# From Salesforce, given a case-sensitive 15-char ID string, translate to its case-safe (18-char) equivalent
# Not required but could be useful later?

function CaseSafeId ( $s ) {
    $a = $raw.ToCharArray()
    $b = 0,0,0
    $c = 'AAA'.ToCharArray()
    for( $i=0; $i -lt 3; $i++ ) {
        for ( $j=0; $j -lt 5; $j++ ) {
            if ( [char]::IsUpper( $a[$i*5+$j] ) ) {
                $b[$i] += ( 1 -shl $j )
                }
            $c[$i] = [char]($b[$i] + 65)
            }
        }
    return $raw + ( -join $c )
}

$raw = '0018000000MFt0S'
CaseSafeId( $raw )

