> **Язык:** **Русский** · [English](../en/publishing.md)
>
> [Обзор](README.md) · [Архитектура и внутреннее устройство](architecture.md) · [Публикация в Maven Central](publishing.md)

# Публикация в Maven Central

Чтобы опубликовать GitHub-репозиторий в Maven Central через новый Central Publisher Portal, в основном нужно выполнить
следующие шаги.

> **NOTE:** `settings.xml` лежит в репозитории. Чтобы случайно не закоммитить секреты, исключите его из отслеживания:
>
> ```bash
> git update-index --skip-worktree settings.xml
> ```

---

### 1. Создайте аккаунт в Central Portal

1. Перейдите на <https://central.sonatype.com> и зарегистрируйтесь (через GitHub или логин/пароль).
   [[Регистрация](https://central.sonatype.org/register/central-portal/#register-to-publish-via-the-central-portal)]

---

### 2. Получите namespace (groupId)

Поскольку код размещён на GitHub, рекомендуемый `groupId`:

- `io.github.<ваш GitHub username>` (для вас, вероятно, `io.github.ogbozoyan`).
  [[Выбор координат](https://central.sonatype.org/publish/requirements/coordinates/#choose-your-coordinates)]

Шаги:

1. В портале откройте меню под вашим username -> **View Namespaces**.
   [[Добавление namespace](https://central.sonatype.org/register/namespace/#adding-a-namespace)]
2. Нажмите **Add Namespace** и введите `io.github.ogbozoyan` (или похожий).
   [[Выбор namespace](https://central.sonatype.org/register/namespace/#choosing-a-namespace)]
3. Для GitHub-based namespace владение подтверждается созданием временного публичного репозитория с именем —
   verification key, который показывает портал (следуйте инструкции «By Code Hosting Services»).
   [[Верификация namespace](https://central.sonatype.org/register/namespace/#verifying-a-namespace)]

Как только namespace получит статус **Verified**, вы сможете публиковать под любым `groupId`, начинающимся с этого
префикса. [[Namespaces vs groupId](https://central.sonatype.org/faq/namespaces-vs-groupids/#how-are-they-related)]

---

### 3. Убедитесь, что POM соответствует требованиям Central

В `pom.xml` минимально необходимо:
[[Обязательные метаданные POM](https://central.sonatype.org/publish/requirements/#required-pom-metadata)]

- Корректные координаты:
  ```xml
  <groupId>io.github.ogbozoyan</groupId>
  <artifactId>report-generator-java</artifactId>
  <version>1.0.0</version> <!-- НЕ должна заканчиваться на -SNAPSHOT -->
  ```
- Packaging (если не дефолтный `jar`):
  ```xml
  <packaging>jar</packaging>
  ```
- Информация о проекте:
  ```xml
  <name>${project.groupId}:${project.artifactId}</name>
  <description>...</description>
  <url>https://github.com/ogbozoyan/report-generator-java</url>
  ```
- Лицензия:
  ```xml
  <licenses>
    <license>
      <name>...</name>
      <url>...</url>
    </license>
  </licenses>
  ```
- Информация о разработчике:
  ```xml
  <developers>
    <developer>
      <name>Your Name</name>
      <email>you@example.com</email>
      <organizationUrl>https://github.com/ogbozoyan</organizationUrl>
    </developer>
  </developers>
  ```
- Информация о SCM:
  ```xml
  <scm>
    <connection>scm:git:git://github.com/ogbozoyan/report-generator-java.git</connection>
    <developerConnection>scm:git:ssh://github.com:ogbozoyan/report-generator-java.git</developerConnection>
    <url>https://github.com/ogbozoyan/report-generator-java</url>
  </scm>
  ```
  [[SCM info](https://central.sonatype.org/publish/requirements/#scm-information)]

---

### 4. Предоставьте javadoc- и sources-jar

Central требует для каждого основного jar: `-sources.jar` и `-javadoc.jar` (они могут содержать реальный или
плейсхолдерный контент, если исходники закрыты).
[[Javadoc & sources](https://central.sonatype.org/publish/requirements/#supply-javadoc-and-sources)][[Closed source allowed](https://central.sonatype.org/faq/closed-source/#can-i-upload-a-closed-source-artifact)]

---

### 5. Настройте GPG-подпись

Все загружаемые файлы должны иметь `.asc`-подписи.
[[GPG requirement](https://central.sonatype.org/publish/requirements/#sign-files-with-gpgpgp)]

1. Установите GnuPG и сгенерируйте ключ: [[GPG setup](https://central.sonatype.org/publish/requirements/gpg/#gpg)]
   ```bash
   gpg --gen-key
   ```
2. Распространите публичный ключ на поддерживаемый keyserver:
   ```bash
   gpg --keyserver keyserver.ubuntu.com --send-keys YOURKEYID
   ```
   [[Distribute key](https://central.sonatype.org/publish/requirements/gpg/#distributing-your-public-key)]

Настройте сборку Maven (например, `maven-gpg-plugin`) для подписи артефактов; документация Central рекомендует
использовать build-инструмент для подписи.
[[Using build tools for signing](https://central.sonatype.org/publish/requirements/gpg/#using-build-tools-for-signing)]

---

### 6. Позвольте Maven-плагину сгенерировать контрольные суммы

Central требует `.md5` и `.sha1` (SHA256/SHA512 опциональны).
[[Checksums requirement](https://central.sonatype.org/publish/requirements/#provide-file-checksums)]

Если вы используете Sonatype `central-publishing-maven-plugin`, он сгенерирует контрольные суммы за вас.
[[Plugin options](https://central.sonatype.org/publish/publish-portal-maven/#plugin-configuration-options)]

---

### 7. Сгенерируйте user token в Central Portal

Публиковать нужно с помощью **user token**, а не пароля от логина.
[[Portal token](https://central.sonatype.org/publish/generate-portal-token/#generating-a-portal-token-for-publishing)]

1. Перейдите на <https://central.sonatype.com/usertoken>.
2. Нажмите **Generate User Token**, задайте имя и срок действия, сохраните пару username/password.

Добавьте её в `~/.m2/settings.xml`:

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
```

[[Maven plugin credentials](https://central.sonatype.org/publish/publish-portal-maven/#publishing-by-using-the-maven-plugin)]

---

### 8. Настройте Maven на публикацию через Central-плагин

Добавьте Sonatype Maven-плагин в `pom.xml`:

```xml
<build>
  <plugins>
    <plugin>
      <groupId>org.sonatype.central</groupId>
      <artifactId>central-publishing-maven-plugin</artifactId>
      <version>0.10.0</version> <!-- или новее -->
      <extensions>true</extensions>
      <configuration>
        <publishingServerId>central</publishingServerId>
        <!-- опционально: автопубликация вместо ручного клика в UI -->
        <!-- <autoPublish>true</autoPublish> -->
      </configuration>
    </plugin>
  </plugins>
</build>
```

[[Maven plugin usage](https://central.sonatype.org/publish/publish-portal-maven/#publishing-by-using-the-maven-plugin)]

Затем:

```bash
mvn clean deploy
```

Плагин:

- соберёт ваш POM, jar, sources, javadoc, подписи, контрольные суммы;
- сформирует bundle;
- загрузит его в Central Publisher Portal.
  [[Publishing flow](https://central.sonatype.org/publish/publish-portal-maven/#publishing)]

По умолчанию он ждёт завершения **validation**; затем вы заходите на портал и нажимаете **Publish**
(если не задано `autoPublish=true`).
[[Publishing & autoPublish](https://central.sonatype.org/publish/publish-portal-maven/#publishing)]

---

### 9. Финальная публикация и неизменяемость

В Central Portal в разделе **Publishing → Deployments** убедитесь, что валидация прошла, и нажмите **Publish** для синка
в Maven Central. [[Portal guide](https://central.sonatype.org/publish/publish-portal-guide/)]

После публикации версию нельзя изменить или удалить — только заместить новой версией.
[[Immutability](https://central.sonatype.org/publish/requirements/immutability/#immutability-of-published-components)]

---

### 10. GitHub workflow

#### Необходимые секреты

Настройте следующие секреты в репозитории GitHub.

Для создания секрета перейдите в: Repository Settings -> Secrets -> Actions -> New repository secret

- `MAVEN_CENTRAL_USERNAME`: сгенерированный Sonatype username
- `MAVEN_CENTRAL_TOKEN`: сгенерированный Sonatype password
- `MAVEN_GPG_PASSPHRASE`: passphrase вашего GPG-ключа
- `MAVEN_GPG_PRIVATE_KEY`: полный приватный GPG-ключ

Экспортируйте приватный ключ командой:

```bash
gpg --armor --export-secret-keys YOUR_KEY_ID
```
