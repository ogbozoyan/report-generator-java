To publish your GitHub repo to Maven Central via the new Central Publisher Portal, you mainly need to:

# NOTE settings.xml BASED IN REPO, IF YOU DON'T WANT TO LEAK ANY SECURE INFORMATION

```bash
git update-index --skip-worktree settings.xml
```

---

### 1. Create a Central Portal account

1. Go to <https://central.sonatype.com> and sign up (GitHub or
   username/password). [[Register account](https://central.sonatype.org/register/central-portal/#register-to-publish-via-the-central-portal)]

---

### 2. Get a namespace (groupId)

Because your code is on GitHub, the recommended groupId is:

- `io.github.<your GitHub username>` (for you, likely
  `io.github.ogbozoyan`). [[Choose coordinates](https://central.sonatype.org/publish/requirements/coordinates/#choose-your-coordinates)]

Steps:

1. In the portal, open the menu under your username -> **View Namespaces
   **. [[Adding namespace](https://central.sonatype.org/register/namespace/#adding-a-namespace)]
2. Click **Add Namespace** and enter `io.github.ogbozoyan` (or
   similar). [[Choosing a namespace](https://central.sonatype.org/register/namespace/#choosing-a-namespace)]
3. For GitHub-based namespaces you prove ownership by creating a temporary public repo named with the verification key
   the portal shows (follow the “By Code Hosting Services”
   instructions). [[Verifying namespace](https://central.sonatype.org/register/namespace/#verifying-a-namespace)]

Once the namespace is **Verified**, you’re allowed to publish under any `groupId` starting with that
prefix. [[Namespaces vs groupId](https://central.sonatype.org/faq/namespaces-vs-groupids/#how-are-they-related)]

---

### 3. Make sure your POM meets Central requirements

In your `pom.xml` you must have at
least: [[Required POM metadata](https://central.sonatype.org/publish/requirements/#required-pom-metadata)]

- Correct coordinates:
  ```xml
  <groupId>io.github.ogbozoyan</groupId>
  <artifactId>report-generator-java</artifactId>
  <version>1.0.0</version> <!-- must NOT end with -SNAPSHOT -->
  ```
- Packaging (if not the default `jar`):
  ```xml
  <packaging>jar</packaging>
  ```
- Project information:
  ```xml
  <name>${project.groupId}:${project.artifactId}</name>
  <description>...</description>
  <url>https://github.com/ogbozoyan/report-generator-java</url>
  ```
- License:
  ```xml
  <licenses>
    <license>
      <name>...</name>
      <url>...</url>
    </license>
  </licenses>
  ```
- Developer info:
  ```xml
  <developers>
    <developer>
      <name>Your Name</name>
      <email>you@example.com</email>
      <organizationUrl>https://github.com/ogbozoyan</organizationUrl>
    </developer>
  </developers>
  ```
- SCM info:
  ```xml
  <scm>
    <connection>scm:git:git://github.com/ogbozoyan/report-generator-java.git</connection>
    <developerConnection>scm:git:ssh://github.com:ogbozoyan/report-generator-java.git</developerConnection>
    <url>https://github.com/ogbozoyan/report-generator-java</url>
  </scm>
  ```[[SCM info](https://central.sonatype.org/publish/requirements/#scm-information)]

---

### 4. Provide javadoc and sources jars

Central requires, for each main jar: a `-sources.jar` and `-javadoc.jar` (they can be real or placeholder content if the
source is
closed). [[Javadoc & sources](https://central.sonatype.org/publish/requirements/#supply-javadoc-and-sources)][[Closed source allowed](https://central.sonatype.org/faq/closed-source/#can-i-upload-a-closed-source-artifact)]

---

### 5. Set up GPG signing

All deployed files need `.asc`
signatures. [[GPG requirement](https://central.sonatype.org/publish/requirements/#sign-files-with-gpgpgp)]

1. Install GnuPG and generate a key: [[GPG setup](https://central.sonatype.org/publish/requirements/gpg/#gpg)]
   ```bash
   gpg --gen-key
   ```
2. Distribute your public key to a supported keyserver:
   ```bash
   gpg --keyserver keyserver.ubuntu.com --send-keys YOURKEYID
   ```[[Distribute key](https://central.sonatype.org/publish/requirements/gpg/#distributing-your-public-key)]

Configure your Maven build (e.g. `maven-gpg-plugin`) to sign artifacts; the Central docs recommend using your build tool
for
signing. [[Using build tools for signing](https://central.sonatype.org/publish/requirements/gpg/#using-build-tools-for-signing)]

---

### 6. Let the Maven plugin generate checksums

Central requires `.md5` and `.sha1` checksums (SHA256/SHA512
optional). [[Checksums requirement](https://central.sonatype.org/publish/requirements/#provide-file-checksums)]

If you use the Sonatype `central-publishing-maven-plugin`, it will generate the checksums for
you. [[Plugin options](https://central.sonatype.org/publish/publish-portal-maven/#plugin-configuration-options)]

---

### 7. Generate a Central Portal user token

You must publish using a **user token**, not your login
password. [[Portal token](https://central.sonatype.org/publish/generate-portal-token/#generating-a-portal-token-for-publishing)]

1. Go to <https://central.sonatype.com/usertoken>.
2. Click **Generate User Token**, set a name and expiration, and save the username/password pair it shows.

Add it to your `~/.m2/settings.xml`:

```xml

<settings>
    <servers>
        <server>
            <id>central</id>
            <username><!-- token username --></username>
            <password><!-- token password --></password>
        </server>
    </servers>
</settings>
        ```[[Maven plugin credentials](https://central.sonatype.org/publish/publish-portal-maven/#publishing-by-using-the-maven-plugin)]

        ---

        ### 8. Configure Maven to publish via the Central plugin

        Add the Sonatype Maven plugin to your `pom.xml`:

        ```xml
<build>
<plugins>
    <plugin>
        <groupId>org.sonatype.central</groupId>
        <artifactId>central-publishing-maven-plugin</artifactId>
        <version>0.10.0</version> <!-- or latest -->
        <extensions>true</extensions>
        <configuration>
            <publishingServerId>central</publishingServerId>
            <!-- optional: auto publish instead of manual click in UI -->
            <!-- <autoPublish>true</autoPublish> -->
        </configuration>
    </plugin>
</plugins>
</build>
```

[[Maven plugin usage](https://central.sonatype.org/publish/publish-portal-maven/#publishing-by-using-the-maven-plugin)]

Then:

```bash
mvn clean deploy
```

The plugin will:

- Collect your POM, jar, sources, javadoc, signatures, checksums
- Build a bundle
- Upload it to the Central Publisher
  Portal [[Publishing flow](https://central.sonatype.org/publish/publish-portal-maven/#publishing)]

By default it waits until **validation** is done; you then go to the portal to click **Publish** (unless
`autoPublish=true`). [[Publishing & autoPublish](https://central.sonatype.org/publish/publish-portal-maven/#publishing)]

---

### 9. Final publish & immutability

In the Central Portal, under **Publishing → Deployments**, check that validation passed and then click **Publish** to
sync to Maven Central. [[Portal guide](https://central.sonatype.org/publish/publish-portal-guide/)]

Once a version is published, it cannot be changed or deleted, only superseded by a new
version. [[Immutability](https://central.sonatype.org/publish/requirements/immutability/#immutability-of-published-components)]

---

If you want, paste your current `pom.xml` and I can point out exactly what to change to satisfy the Central
requirements.

### 10. Github workflow

#### Required secrets configuration

You need to configure the following secrets in your GitHub repository.

To create a secret, navigate to: Repository Settings -> Secrets -> Actions -> New repository secret

MAVEN_CENTRAL_USERNAME: your Sonatype-generated username
MAVEN_CENTRAL_TOKEN: your Sonatype-generated password
MAVEN_GPG_PASSPHRASE: the passphrase for your GPG key
MAVEN_GPG_PRIVATE_KEY: your complete GPG private key

- Export it using: gpg --armor --export-secret-keys YOUR_KEY_ID