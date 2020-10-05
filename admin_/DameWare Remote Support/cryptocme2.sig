<?rsa version="1.0" encoding="utf-8"?>
<Configuration>
	<Product Id="Crypto-C ME">
		<Version>CRYPTO-C ME 3.0.0.0</Version>
		<ReleaseDate>""</ReleaseDate>
		<ExpDate>""</ExpDate>
		<Copyright>
			Copyright (C) RSA
		</Copyright>
		<Library Id="master">cryptocme2</Library>
	</Product>
	<Runtime Id="runtime">
		<LoadOrder>
			<Library Id="ccme_base">ccme_base</Library>
			<Library Id="ccme_ecc">ccme_ecc</Library>
			<Library Id="ccme_eccaccel">ccme_eccaccel</Library>
			<Library Id="ccme_eccnistaccel">ccme_eccnistaccel</Library>
		</LoadOrder>
		<StartupConfig>
			<SelfTest>OnLoad</SelfTest>
		</StartupConfig>
	</Runtime>
	<Signature URI="#ccme_base" Algorithm="FIPS140_INTEGRITY">MC0CFQC07Sgcr3KzdCZ37Q13i0pyGMGK4wIUOGufSwQFIXsxdljio4dk3kU6tNY=</Signature>
	<Signature URI="#ccme_ecc" Algorithm="FIPS140_INTEGRITY">MCwCFHxOgbakOlht+00QjMHGWI8ioem3AhQZJd4rW/n8CK5uJu4EdKZXRyIkQA==</Signature>
	<Signature URI="#ccme_eccaccel" Algorithm="FIPS140_INTEGRITY">MC0CFQC0U3qt1awwPlAkk7u2tKdiJPIlJwIUYEKwEipZWBgZ/7nHN18RfiW6Dxo=</Signature>
	<Signature URI="#ccme_eccnistaccel" Algorithm="FIPS140_INTEGRITY">MCsCFDWSOM3hcXKKbarLn6CawnJvpUF5AhPySQnNgOIguXoVY91KtccUqy3m</Signature>
	<Signature URI="#master" Algorithm="FIPS140_INTEGRITY">MCwCFFSXDsWDJ1P7GG9UGOj60YSbtrZ6AhRVgqFFy4JoV1gXAM5NdQ4rIG+/XQ==</Signature>
	<Signature URI="#Crypto-C ME" Algorithm="FIPS140_INTEGRITY">MC0CFHwWAJz3N4b4KWNidFproESOD4UtAhUAw3GJYx3mEtpIxXWIF5zZ6RQSbzY=</Signature>
	<Signature URI="#runtime" Algorithm="FIPS140_INTEGRITY">MCwCFDeYEE66gkaRZ8waIjWMsPrFzjJkAhQ6+sX6D2mFlChxPtZXyiCiwiRe4w==</Signature>
</Configuration>

