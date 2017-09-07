package i18n.xlsx;

/**
 * Key object to serve the translation map. <br/>
 * Supports a case-insensitive search.
 */
public final class LngKey {

    private String language;
    private String phrase;

    /**
     * Default constructor
     */
    public LngKey() {
        super();
    }

    /**
     * Constructor with all fields.
     * 
     * @param language
     * @param phrase
     */
    public LngKey(final String language, final String phrase) {
        super();
        this.language = language;
        this.phrase = phrase;
    }

    /**
     * @return the language
     */
    public String getLanguage() {
        return this.language;
    }

    /**
     * @param language
     *            the language to set
     */
    public void setLanguage(final String language) {
        this.language = language;
    }

    /**
     * @return the phrase
     */
    public String getPhrase() {
        return this.phrase;
    }

    /**
     * @param phrase
     *            the phrase to set
     */
    public void setPhrase(final String phrase) {
        this.phrase = phrase;
    }

    /**
     * @see java.lang.Object#hashCode()
     */
    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = (prime * result) + ((this.language == null) ? 0 : this.language.toUpperCase().hashCode());
        result = (prime * result) + ((this.phrase == null) ? 0 : this.phrase.toUpperCase().hashCode());
        return result;
    }

    /**
     * @see java.lang.Object#equals(Object)
     */
    @Override
    public boolean equals(final Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (this.getClass() != obj.getClass()) {
            return false;
        }
        final LngKey other = (LngKey) obj;
        if (this.language == null) {
            if (other.language != null) {
                return false;
            }
        } else if (!this.language.equalsIgnoreCase(other.language)) {
            return false;
        }
        if (this.phrase == null) {
            if (other.phrase != null) {
                return false;
            }
        } else if (!this.phrase.equalsIgnoreCase(other.phrase)) {
            return false;
        }
        return true;
    }
}
